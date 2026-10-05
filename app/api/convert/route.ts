import { NextResponse } from 'next/server';
import { GoogleGenAI } from '@google/genai';
import pptxgen from 'pptxgenjs';

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

async function generateWithRetry(fileBytes: Buffer, mimeType: string, prompt: string, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await ai.models.generateContent({
        model: 'gemini-3.8-flash',
        contents: [
          {
            inlineData: {
              mimeType: mimeType,
              data: fileBytes.toString('base64'),
            },
          },
          prompt,
        ],
      });
      return response.text || '';
    } catch (error: any) {
      console.warn(`Tentativo \({attempt}/\){maxRetries} fallito:`, error.message);
      if (attempt === maxRetries) throw error;
      await new Promise((resolve) => setTimeout(resolve, attempt * 1500));
    }
  }
  throw new Error('Superato il limite massimo di tentativi con Gemini.');
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

    const prompt = `
Sei un esperto di finanza e corporate storytelling. Analizza il documento allegato e suddividi i contenuti in esattamente 6-7 sezioni/slide. 
Per ogni slide restituisci un titolo chiaro e una lista di punti elenco dettagliati e professionali.
Usa rigorosamente questo formato esatto per ogni slide:
=== SLIDE N: [Titolo della Slide] ===
- [Punto elenco 1 con dettagli e metriche]
- [Punto elenco 2]
- [Punto elenco 3]

Contesto di riferimento:
- Obiettivo: ${objective}
- Settore: ${sector}
- Note: ${userAnswers || 'Nessuna'}
    `;

    const rawText = await generateWithRetry(fileBytes, file.type || 'application/pdf', prompt);

    const pptx = new pptxgen();
    pptx.layout = 'LAYOUT_16x9';

    const slideChunks = rawText.split('=== SLIDE').filter(Boolean);

    if (slideChunks.length === 0) {
      const slide = pptx.addSlide();
      slide.addText(rawText, { x: 0.8, y: 0.8, w: '85%', h: '80%', fontSize: 14, color: '333333' });
    } else {
      for (const chunk of slideChunks) {
        const lines = chunk.split('\n').map((l) => l.trim()).filter(Boolean);
        const headerLine = lines[0].replace(/^[:=]+/, '').trim();
        const bulletPoints = lines.slice(1).map((l) => l.replace(/^[-*]\s*/, ''));

        const slide = pptx.addSlide();

        slide.addText(headerLine, {
          x: 0.8,
          y: 0.6,
          w: '85%',
          h: 0.8,
          fontSize: 22,
          bold: true,
          color: '4A154B',
        });

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
    }

    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });

    // Cast esplicito per soddisfare il controllo di tipo BodyInit in Next.js
    return new NextResponse(pptxBuffer as unknown as BodyInit, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="presentazione-docudecky.pptx"',
      },
    });

  } catch (error: any) {
    console.error('Errore API Generazione PPTX:', error);
    return NextResponse.json(
      { message: error.message || 'Errore interno durante la generazione della presentazione.' },
      { status: 500 }
    );
  }
}
