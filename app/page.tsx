'use client';

import React, { useState } from 'react';
import { Upload, Sparkles, Loader2 } from 'lucide-react';
import pptxgen from 'pptxgenjs';

export default function Home() {
  const [file, setFile] = useState(null);
  const [sector, setSector] = useState('Business / Aziendale');
  const [objective, setObjective] = useState('Presentazione PPTX');
  const [loading, setLoading] = useState(false);
  const [statusText, setStatusText] = useState('');

  const handleConvert = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) return;

    setLoading(true);
    setStatusText('Lettura del documento in corso...');

    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('sector', sector);
      formData.append('objective', objective);

      setStatusText('Analisi IA e generazione slide...');
      const res = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
      });

      if (!res.ok) {
        const errorData = await res.json().catch(() => ({ message: 'Errore durante la conversione' }));
        throw new Error(errorData.message || 'Errore durante la conversione');
      }

      const resJson = await res.json();
      const data = resJson.data;

      setStatusText('Download presentazione in corso...');

      const pptx = new pptxgen();
      pptx.layout = 'LAYOUT_16x9';

      const titleSlide = pptx.addSlide();
      titleSlide.background = { color: 'F8FAFC' };

      titleSlide.addText(data.title || 'Presentazione', {
        x: 1.0,
        y: 2.0,
        w: 11.3,
        h: 1.5,
        fontSize: 36,
        fontFace: 'Arial',
        bold: true,
        color: '0F172A',
        align: 'left',
      });

      if (data.summary) {
        titleSlide.addText(data.summary, {
          x: 1.0,
          y: 3.8,
          w: 11.3,
          h: 2.0,
          fontSize: 16,
          fontFace: 'Arial',
          color: '475569',
          align: 'left',
        });
      }

      if (data.slides && data.slides.length > 0) {
        data.slides.forEach((item: any) => {
          const slide = pptx.addSlide();
          slide.background = { color: 'FFFFFF' };

          slide.addText(item.slideTitle || 'Nuova Slide', {
            x: 0.8,
            y: 0.8,
            w: 11.7,
            h: 0.8,
            fontSize: 24,
            fontFace: 'Arial',
            bold: true,
            color: '1E293B',
          });

          if (item.bulletPoints && item.bulletPoints.length > 0) {
            const formattedBullets = item.bulletPoints.map((point: string) => ({
              text: point,
              options: {
                bullet: true,
                fontSize: 16,
                fontFace: 'Arial',
                color: '334155',
                spaceAfter: 12,
              },
            }));

            slide.addText(formattedBullets, {
              x: 0.8,
              y: 1.8,
              w: 11.7,
              h: 4.8,
              align: 'left',
              valign: 'top',
            });
          }
        });
      }

      const nameParts = file.name.split('.');
      if (nameParts.length > 1) {
        nameParts.pop();
      }
      const baseName = nameParts.join('.') || 'documento';
      
      await pptx.writeFile({ fileName: `${baseName}-presentazione.pptx` });

    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Errore sconosciuto';
      alert(`Errore: ${msg}`);
    } finally {
      setLoading(false);
      setStatusText('');
    }
  };

  return
