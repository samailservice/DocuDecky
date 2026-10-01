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
    setStatusText('Elaborazione IA in corso...');

    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('sector', sector);
      formData.append('objective', objective);

      const res = await fetch('/api/convert', { method: 'POST', body: formData });
      const json = await res.json();

      if (!res.ok) throw new Error(json.message || 'Errore di conversione');

      setStatusText('Creazione presentazione...');
      const pptx = new pptxgen();
      pptx.layout = 'LAYOUT_16x9';

      const data = json.data;
      const tSlide = pptx.addSlide();
      tSlide.addText(data.title || 'Presentazione', { x: 1, y: 2, w: 11, h: 1.5, fontSize: 36, bold: true });
      if (data.summary) tSlide.addText(data.summary, { x: 1, y: 3.8, w: 11, h: 2, fontSize: 16 });

      if (data.slides) {
        data.slides.forEach((s: any) => {
          const slide = pptx.addSlide();
          slide.addText(s.slideTitle || 'Slide', { x: 0.8, y: 0.8, w: 11.7, h: 0.8, fontSize: 24, bold: true });
          if (s.bulletPoints) {
            slide.addText(s.bulletPoints.map((p: string) => ({ text: p, options: { bullet: true, fontSize: 16 } })), { x: 0.8, y: 1.8, w: 11.7, h: 4.8 });
          }
        });
      }

      const baseName = file.name.substring(0, file.name.lastIndexOf('.')) || 'documento';
      await pptx.writeFile({ fileName: `${baseName}-presentazione.pptx` });
    } catch (err: any) {
      alert(`Errore: ${err.message}`);
    } finally {
      setLoading(false);
      setStatusText('');
    }
  };

  return (
