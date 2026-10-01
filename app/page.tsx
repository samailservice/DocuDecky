'use client';

import React, { useState } from 'react';
import { Upload, Sparkles, Loader2 } from 'lucide-react';

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [sector, setSector] = useState('Business / Aziendale');
  const [objective, setObjective] = useState('Presentazione PPTX');
  const [loading, setLoading] = useState(false);
  const [statusText, setStatusText] = useState('');

  const handleConvert = async (e: React.FormEvent<HTMLFormElement>) => {
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

  return <main className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6">
    <div className="max-w-xl w-full bg-white rounded-2xl shadow-xl p-8 border border-slate-100">
      <div className="text-center mb-8">
        <div className="inline-flex items-center justify-center w-16 h-16 rounded-full bg-indigo-50 text-indigo-600 mb-4">
          <Sparkles className="w-8 h-8" />
        </div>
        <h1 className="text-2xl font-bold text-slate-900">DocuDecky AI</h1>
        <p className="text-slate-500 text-sm mt-1">
          Trascina un documento e trasformalo in una presentazione professionale.
        </p>
      </div>

      <form onSubmit={handleConvert} className="space-y-6">
        <div>
          <label className="block text-sm font-medium text-slate-700 mb-2">
            Carica Documento (PDF, TXT, DOCX)
          </label>
          <div className="flex items-center justify-center w-full">
            <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-slate-300 border-dashed rounded-xl cursor-pointer bg-slate-50 hover:bg-slate-100 transition">
              <div className="flex flex-col items-center justify-center pt-5 pb-6 px-4 text-center">
                <Upload className="w-8 h-8 mb-2 text-slate-400" />
                <p className="text-sm text-slate-600">
                  {file ? (
                    <span className="font-semibold text-indigo-600">{file.name}</span>
                  ) : (
                    'Clicca o trascina il file qui'
                  )}
                </p>
              </div>
              <input
                type="file"
                className="hidden"
                onChange={(e) => e.target.files && setFile(e.target.files[0])}
                accept=".pdf,.txt,.docx,.doc"
              />
            </label>
          </div>
        </div>

        <div>
          <label className="block text-sm font-medium text-slate-700 mb-2">Settore</label>
          <select
            value={sector}
            onChange={(e) => setSector(e.target.value)}
            className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:outline-none text-slate-700 bg-white"
          >
            <option value="Business / Aziendale">Business / Aziendale</option>
            <option value="Educativo / Accademico">Educativo / Accademico</option>
            <option value="Tecnologico / Startup">Tecnologico / Startup</option>
            <option value="Creativo / Marketing">Creativo / Marketing</option>
          </select>
        </div>

        <div>
          <label className="block text-sm font-medium text-slate-700 mb-2">Obiettivo</label>
          <select
            value={objective}
            onChange={(e) => setObjective(e.target.value)}
            className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:outline-none text-slate-700 bg-white"
          >
            <option value="Presentazione PPTX">Presentazione PPTX</option>
            <option value="Pitch Deck">Pitch Deck</option>
            <option value="Riassunto Esecutivo">Riassunto Esecutivo</option>
          </select>
        </div>

        <button
          type="submit"
          disabled={!file || loading}
          className="w-full bg-indigo-600 text-white font-medium py-3 px-4 rounded-xl hover:bg-indigo-700 transition disabled:opacity-50 disabled:cursor-not-allowed flex items-center justify-center gap-2 shadow-lg shadow-indigo-100"
        >
          {loading ? (
            <>
              <Loader2 className="w-5 h-5 animate-spin" />
              <span>{statusText}</span>
            </>
          ) : (
            <>
              <Sparkles className="w-5 h-5" />
              <span>Genera Presentazione</span>
            </>
          )}
        </button>
      </form>
    </div>
  </main>;
}
