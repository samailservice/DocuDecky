'use client';

import React, { useState } from 'react';

export default function Page() {
  const [loading, setLoading] = useState(false);
  const [questions, setQuestions] = useState([]);
  const [answers, setAnswers] = useState('');
  const [file, setFile] = useState(null);
  const [sector, setSector] = useState('Business / Aziendale');
  const [objective, setObjective] = useState('Presentazione PPTX');
  const [statusText, setStatusText] = useState('');

  const handleConvert = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setStatusText('Elaborazione in corso...');

    const formData = new FormData();
    if (file) {
      formData.append('file', file);
    }
    formData.append('sector', sector);
    formData.append('objective', objective);
    if (answers) {
      formData.append('userAnswers', answers);
    }

    try {
      const res = await fetch('/api/convert', {
        method: 'POST',
        body: formData,
      });

      const contentType = res.headers.get('content-type');
      if (contentType && contentType.includes('application/json')) {
        const data = await res.json();
        if (data.needsInput) {
          setQuestions(data.questions);
          setLoading(false);
          setStatusText('');
          return;
        }
      }

      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        alert(errData.message || 'Errore durante la generazione');
        setLoading(false);
        setStatusText('');
        return;
      }

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = res.headers.get('content-disposition')?.split('filename="')[1]?.replace('"', '') || 'output.pptx';
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (err: any) {
      console.error(err);
      alert('Errore di connessione al server');
    } finally {
      setLoading(false);
      setStatusText('');
    }
  };

  return (
    <main className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-6">
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
    </main>
  );
}
