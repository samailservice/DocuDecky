'use client';

import React, { useState } from 'react';
import { Upload, Sparkles, Loader2 } from 'lucide-react';

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
        const errorData = await res.json();
        throw new Error(errorData.message || 'Errore durante la conversione');
      }

      setStatusText('Download presentazione in corso...');
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `\({file.name.replace(/\.[^/.]+\)/, '')}-presentazione.pptx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err: any) {
      alert(`Errore: ${err.message}`);
    } finally {
      setLoading(false);
      setStatusText('');
    }
  };

  return (
