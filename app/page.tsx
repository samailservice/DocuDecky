'use client';

import React, { useState } from 'react';

export default function Page() {
  const [loading, setLoading] = useState(false);
  const [questions, setQuestions] = useState<string[]>([]);
  const [answers, setAnswers] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const [sector, setSector] = useState('Business / Aziendale');
  const [objective, setObjective] = useState('presentazione_pptx');
  const [statusText, setStatusText] = useState('');
  const [successMessage, setSuccessMessage] = useState('');

  const handleConvert = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setStatusText('Analisi del documento e contestualizzazione AI in corso...');
    setSuccessMessage('');

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
          setQuestions(data.questions || []);
          setLoading(false);
          setStatusText('');
          return;
        }

        if (data.success) {
          setLoading(false);
          setStatusText('');
          setSuccessMessage('Elaborazione completata con successo!');
          if (data.data && objective === 'documento_docx') {
            const blob = new Blob([data.data], { type: 'text/plain;charset=utf-8' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = data.filename || 'report-strategico.txt';
            document.body.appendChild(a);
            a.click();
            a.remove();
          }
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

      // Ricezione file binario (es. .pptx)
      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const disposition = res.headers.get('content-disposition');
      let filename = 'presentazione.pptx';
      if (disposition && disposition.includes('filename="')) {
        filename = disposition.split('filename="')[1].replace('"', '');
      }
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setSuccessMessage('File scaricato con successo!');
    } catch (err: any) {
      console.error(err);
      alert('Errore di connessione al server');
    } finally {
      setLoading(false);
      setStatusText('');
    }
  };

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100 flex flex-col items-center justify-center p-6 selection:bg-sky-500 selection:text-white">
      <div className="w-full max-w-4xl bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl p-8 md:p-12">
        
        {/* Header */}
        <div className="text-center mb-10">
          <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-sky-500/10 text-sky-400 text-xs font-semibold uppercase tracking-wider mb-4 border border-sky-500/20">
            <i className="fa-solid fa-wand-magic-sparkles"></i> DocuDecky AI Engine
          </div>
          <h1 className="text-3xl md:text-5xl font-extrabold tracking-tight mb-4">
            Converti i tuoi Report in <span className="text-transparent bg-clip-text bg-gradient-to-r from-sky-400 to-indigo-500">Deck Professionali</span>
          </h1>
          <p className="text-slate-400 text-base md:text-lg max-w-2xl mx-auto">
            Carica un documento (PDF, Word, Bilancio) e seleziona l'obiettivo desiderato per ottenere contenuti su misura.
          </p>
        </div>

        {/* Form principale */}
        <form onSubmit={handleConvert} className="space-y-6">
          
          {/* Upload file */}
          <div className="border-2 border-dashed border-slate-700 hover:border-sky-500 transition-all rounded-xl p-8 text-center bg-slate-950/50 group cursor-pointer relative">
            <input 
              type="file" 
              onChange={(e) => setFile(e.target.files?.[0] || null)}
              className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
            />
            <div className="flex flex-col items-center justify-center pointer-events-none">
              <div className="w-16 h-16 rounded-full bg-sky-500/10 flex items-center justify-center text-sky-400 mb-4 group-hover:scale-110 transition-transform">
                <i className="fa-solid fa-cloud-arrow-up text-2xl"></i>
              </div>
              <p className="text-lg font-semibold text-slate-200 mb-1">
                {file ? file.name : "Trascina qui il tuo file o clicca per cercarlo"}
              </p>
              <p className="text-sm text-slate-400">
                Supporta PDF, DOCX, TXT (Max 25MB)
              </p>
            </div>
          </div>

          {/* Griglia Selezione Settore e Obiettivo */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-sm font-semibold text-slate-300 mb-2">
                <i className="fa-solid fa-briefcase mr-2 text-sky-400"></i> Settore di Riferimento
              </label>
              <select 
                value={sector} 
                onChange={(e) => setSector(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-3 text-slate-200 focus:outline-none focus:border-sky-500 transition-colors"
              >
                <option value="Business / Aziendale">Business / Aziendale</option>
                <option value="Finanza & Banking">Finanza & Banking</option>
                <option value="ESG & Sostenibilità">ESG & Sostenibilità</option>
                <option value="Real Estate & Immobiliare">Real Estate & Immobiliare</option>
                <option value="Startup & VC / Tech">Startup & VC / Tech</option>
              </select>
            </div>

            <div>
              <label className="block text-sm font-semibold text-slate-300 mb-2">
                <i className="fa-solid fa-bullseye mr-2 text-sky-400"></i> Obiettivo di Output
              </label>
              <select 
                value={objective} 
                onChange={(e) => setObjective(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-3 text-slate-200 focus:outline-none focus:border-sky-500 transition-colors"
              >
                <option value="presentazione_pptx">Presentazione PowerPoint (.pptx)</option>
                <option value="google_slides">Compatibile Google Slides (.pptx)</option>
                <option value="pitch_presentazione">Pitch Deck Persuasivo</option>
                <option value="documento_docx">Documento / Report Testuale Strutturato</option>
              </select>
            </div>
          </div>

          {/* Se l'IA richiede input aggiuntivi */}
          {questions.length > 0 && (
            <div className="bg-sky-500/10 border border-sky-500/30 rounded-xl p-6 space-y-4">
              <div className="flex items-center gap-2 text-sky-400 font-semibold">
                <i className="fa-solid fa-circle-question"></i> L'IA richiede qualche chiarimento prima di procedere:
              </div>
              <ul className="list-disc list-inside text-sm text-slate-300 space-y-1">
                {questions.map((q, idx) => (
                  <li key={idx}>{q}</li>
                ))}
              </ul>
              <div>
                <label className="block text-sm font-medium text-slate-300 mb-1">Le tue risposte o indicazioni aggiuntive:</label>
                <textarea 
                  value={answers}
                  onChange={(e) => setAnswers(e.target.value)}
                  placeholder="Es. Il focus principale deve essere sui dati finanziari del Q3..."
                  rows={3}
                  className="w-full bg-slate-950 border border-slate-800 rounded-lg p-3 text-slate-200 focus:outline-none focus:border-sky-500"
                />
              </div>
            </div>
          )}

          {/* Bottone di Invio / Azione */}
          <button 
            type="submit"
            disabled={loading || !file}
            className="w-full bg-gradient-to-r from-sky-500 to-indigo-600 hover:from-sky-400 hover:to-indigo-500 disabled:opacity-50 disabled:cursor-not-allowed text-white font-bold py-4 rounded-xl shadow-lg transition-all flex items-center justify-center gap-3 text-lg"
          >
            {loading ? (
              <>
                <i className="fa-solid fa-spinner fa-spin"></i>
                <span>{statusText || 'Elaborazione in corso...'}</span>
              </>
            ) : (
              <>
                <i className="fa-solid fa-wand-magic-sparkles"></i>
                <span>Genera Output Richiesto</span>
              </>
            )}
          </button>

          {successMessage && (
            <div className="p-4 bg-emerald-500/10 border border-emerald-500/30 rounded-xl text-emerald-400 text-center font-medium">
              <i className="fa-solid fa-circle-check mr-2"></i> {successMessage}
            </div>
          )}

        </form>

      </div>
    </main>
  );
}
