'use client';

import React, { useState } from 'react';

export default function Page() {
  const [loading, setLoading] = useState(false);
  const [questions, setQuestions] = useState<string[]>([]);
  const [answers, setAnswers] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const [sector, setSector] = useState('Business / Aziendale');
  const [objective, setObjective] = useState('Presentazione PPTX');
  const [statusText, setStatusText] = useState('');

  const handleConvert = async (e: React.FormEvent<HTMLFormElement>) => {
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
    <main className="min-h-screen bg-gray-50 flex flex-col items-center justify-center p-6">
      <div className="max-w-xl w-full bg-white p-8 rounded-xl shadow-md">
        <h1 className="text-2xl font-bold mb-4 text-gray-800 flex items-center gap-2">
          ✨ DocuDecky AI
        </h1>
        
        {questions.length > 0 && (
          <div className="bg-amber-50 border border-amber-200 p-4 rounded-lg mb-4">
            <h2 className="font-semibold text-amber-800 mb-2">L'IA richiede chiarimenti:</h2>
            <ul className="list-disc pl-5 mb-4 text-sm text-amber-700">
              {questions.map((q, idx) => (
                <li key={idx}>{q}</li>
              ))}
            </ul>
            <textarea
              className="w-full p-2 border rounded mb-2 text-sm text-gray-700"
              rows={3}
              placeholder="Scrivi qui le tue risposte..."
              value={answers}
              onChange={(e) => setAnswers(e.target.value)}
            />
            <button
              type="button"
              onClick={() => {
                const form = document.getElementById('convert-form') as HTMLFormElement;
                if (form) form.requestSubmit();
              }}
              className="bg-purple-600 text-white px-4 py-2 rounded text-sm hover:bg-purple-700"
            >
              Invia risposte e continua
            </button>
          </div>
        )}

        <form id="convert-form" onSubmit={handleConvert} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1 flex items-center gap-1">
              📁 Carica Documento (PDF, TXT, DOCX)
            </label>
            <input 
              type="file" 
              required 
              onChange={(e) => setFile(e.target.files?.[0] || null)}
              className="w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-md file:border-0 file:text-sm file:font-semibold file:bg-purple-50 file:text-purple-700 hover:file:bg-purple-100" 
            />
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Settore</label>
            <select 
              value={sector}
              onChange={(e) => setSector(e.target.value)}
              className="w-full p-2 border rounded-md text-sm text-gray-800 bg-white"
            >
              <option value="Business / Aziendale">Business / Aziendale</option>
              <option value="Tecnologia">Tecnologia</option>
              <option value="Sanità">Sanità</option>
              <option value="Energia">Energia</option>
            </select>
          </div>

          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Obiettivo</label>
            <input 
              type="text" 
              value={objective}
              onChange={(e) => setObjective(e.target.value)}
              className="w-full p-2 border rounded-md text-sm text-gray-800" 
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-purple-600 text-white py-2 px-4 rounded-md font-medium hover:bg-purple-700 transition disabled:opacity-50 flex items-center justify-center gap-2"
          >
            {loading ? (
              <>
                <span className="animate-spin">⏳</span>
                {statusText || 'Analisi IA e generazione in corso...'}
              </>
            ) : (
              <>✨ Genera Presentazione</>
            )}
          </button>
        </form>
      </div>
    </main>
  );
}
