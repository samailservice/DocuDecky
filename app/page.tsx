'use client';

import React, { useState } from 'react';

import { Upload, FileText, Sparkles, Send, Bot } from 'lucide-react';

export default function Home() {

  const [file, setFile] = useState<File | null>(null);

  const [sector, setSector] = useState('university');

  const [goal, setGoal] = useState('presentation');

  const [loading, setLoading] = useState(false);

  const [result, setResult] = useState<any>(null);

  const [chatOpen, setChatOpen] = useState(false);

  const [messages, setMessages] = useState([

    { role: 'bot', text: 'Ciao! Sono l’assistente AI di DocuDecky. Come posso aiutarti?' }

  ]);

  const [inputMsg, setInputMsg] = useState('');

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {

    if (e.target.files && e.target.files[0]) {

      setFile(e.target.files[0]);

    }

  };

  const handleConvert = async () => {

    if (!file) return alert('Seleziona prima un file!');

    setLoading(true);

    const formData = new FormData();

    formData.append('file', file);

    formData.append('sector', sector);

    formData.append('goal', goal);

    try {

      const res = await fetch('/api/convert', { method: 'POST', body: formData });

      const data = await res.json();

      if (data.success) {

        setResult(data.data);

      } else {

        alert('Errore: ' + data.error);

      }

    } catch (err) {

      alert('Si è verificato un errore durante la conversione.');

    } finally {

      setLoading(false);

    }

  };

  const downloadPPTX = async () => {

    if (!result) return;

    try {

      const pptxgen = (await import('pptxgenjs')).default;

      const pres = new pptxgen();

      

      const titleSlide = pres.addSlide();

      titleSlide.addText(result.title, { x: 1, y: 2, fontSize: 32, bold: true, color: '003366' });

      titleSlide.addText(result.summary, { x: 1, y: 3.5, fontSize: 18, color: '666666' });

      if (result.slides && Array.isArray(result.slides)) {

        result.slides.forEach((s: any) => {

          const slide = pres.addSlide();

          slide.addText(s.slideTitle, { x: 0.8, y: 0.8, fontSize: 24, bold: true, color: '003366' });

          if (s.bulletPoints && Array.isArray(s.bulletPoints)) {

            slide.addText(s.bulletPoints.map((bp: string) => `• ${bp}`).join('\n'), {

              x: 0.8, y: 2.0, fontSize: 16, color: '333333', lineSpacing: 28

            });

          }

        });

      }

      await pres.writeFile({ fileName: `${result.title || 'DocuDecky'}.pptx` });

    } catch (e) {

      console.error(e);

      alert('Errore durante la generazione del file PPTX');

    }

  };

  return (

    <main className="flex min-h-screen flex-col items-center justify-between p-8 bg-slate-900 text-white">

      <div className="w-full max-w-4xl mx-auto space-y-8">

        <header className="flex justify-between items-center border-b border-slate-800 pb-4">

          <div className="flex items-center space-x-2">

            <Sparkles className="w-8 h-8 text-indigo-400" />

            <h1 className="text-2xl font-bold tracking-tight">DocuDecky</h1>

          </div>

          <span className="text-xs bg-indigo-500/10 text-indigo-400 px-3 py-1 rounded-full border border-indigo-500/20">Next.js 15 & AI</span>

        </header>

        <div className="bg-slate-800/50 border border-slate-700/60 rounded-2xl p-6 shadow-xl space-y-6">

          <div className="space-y-2">

            <label className="text-sm font-medium text-slate-300">Carica il tuo documento</label>

            <div className="flex items-center justify-center w-full">

              <label className="flex flex-col items-center justify-center w-full h-32 border-2 border-slate-600 border-dashed rounded-xl cursor-pointer bg-slate-800/80 hover:bg-slate-700/50 transition">

                <div className="flex flex-col items-center justify-center pt-5 pb-6">

                  <Upload className="w-8 h-8 mb-2 text-slate-400" />

                  <p className="text-sm text-slate-300">

                    <span className="font-semibold">{file ? file.name : "Clicca per caricare"}</span> o trascina il file

                  </p>

                  <p className="text-xs text-slate-500">PDF, DOCX o TXT</p>

                </div>

                <input type="file" className="hidden" onChange={handleFileChange} />

              </label>

            </div>

          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">

            <div className="space-y-2">

              <label className="text-sm font-medium text-slate-300">Settore</label>

              <select 

                value={sector} 

                onChange={(e) => setSector(e.target.value)}

                className="w-full bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white focus:ring-2 focus:ring-indigo-500"

              >

                <option value="university">Università / Accademico</option>

                <option value="business">Business / Aziendale</option>

                <option value="creative">Creativo / Pitch</option>

              </select>

            </div>

            <div className="space-y-2">

              <label className="text-sm font-medium text-slate-300">Obiettivo</label>

              <select 

                value={goal} 

                onChange={(e) => setGoal(e.target.value)}

                className="w-full bg-slate-900 border border-slate-700 rounded-lg p-2.5 text-sm text-white focus:ring-2 focus:ring-indigo-500"

              >

                <option value="presentation">Presentazione PPTX</option>

                <option value="summary">Riassunto / Dispensa</option>

              </select>

            </div>

          </div>

          <button

            onClick={handleConvert}

            disabled={loading}

            className="w-full bg-indigo-600 hover:bg-indigo-500 text-white font-medium py-3 rounded-xl transition flex items-center justify-center space-x-2 disabled:opacity-50"

          >

            <Sparkles className="w-5 h-5" />

            <span>{loading ? 'Elaborazione con IA in corso...' : 'Genera con AI'}</span>

          </button>

        </div>

        {result && (

          <div className="bg-slate-800/50 border border-slate-700/60 rounded-2xl p-6 space-y-4">

            <h2 className="text-xl font-bold text-indigo-300">{result.title}</h2>

            <p className="text-slate-300 text-sm">{result.summary}</p>

            <button

              onClick={downloadPPTX}

              className="bg-emerald-600 hover:bg-emerald-500 text-white px-4 py-2 rounded-lg text-sm font-medium transition flex items-center space-x-2"

            >

              <FileText className="w-4 h-4" />

              <span>Scarica Presentazione PPTX</span>

            </button>

          </div>

        )}

      </div>

    </main>

  );

}
