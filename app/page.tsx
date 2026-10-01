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
        const errorData = await res.json().catch(() => ({ message: 'Errore durante la conversione' }));
        throw new Error(errorData.message || 'Errore durante la conversione');
      }

      setStatusText('Download presentazione in corso...');
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;

      const nameParts = file.name.split('.');
      if (nameParts.length > 1) {
        nameParts.pop();
      }
      const baseName = nameParts.join('.') || 'documento';
      a.download = `${baseName}-presentazione.pptx`;

      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : 'Errore sconosciuto';
      alert(`Errore: ${msg}`);
    } finally {
      setLoading(false);
      setStatusText('');
    }
  };

  return (
