'use client';

import React, { useState } from 'react';
import { Upload, FileText, Sparkles, Send, Bot } from 'lucide-react';
import pptxgen from 'pptxgenjs';

export default function Home() {
  const [file, setFile] = useState(null);
  const [sector, setSector] = useState('university');
  const [goal, setGoal] = useState('presentation');
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);

  const [chatOpen, setChatOpen] = useState(false);
  const [messages, setMessages] = useState([
    { role: 'bot', text: 'Ciao! Sono l’assistente AI di DocuDecky. Come posso aiutarti?' }
  ]);
  const [inputMsg, setInputMsg] = useState('');

  const handleFileChange = (e: React.ChangeEvent) => {
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

  const downloadPPTX = () => {
    if (!result) return;
    const pres = new pptxgen();
    
    const titleSlide = pres.addSlide();
    titleSlide.addText(result.title, { x: 1, y: 2, fontSize: 32, bold: true, color: '003366' });
    titleSlide.addText(result.summary, { x: 1, y: 3.5, fontSize: 18, color: '666666' });

    result.slides.forEach((s: any) => {
      const slide = pres.addSlide();
      slide.addText(s.slideTitle, { x: 0.8, y: 0.8, fontSize: 24, bold: true, color: '003366' });
      slide.addText(s.bulletPoints.map((bp: string) => `• ${bp}`).join('\n'), {
        x: 0.8, y: 2.0, fontSize: 16, color: '333333', lineSpacing: 28
      });
    });

    pres.writeFile({ fileName: `${result.title || 'DocuDecky'}.pptx` });
  };

  const handleSendMessage = () => {
    if (!inputMsg.trim()) return;
    const newMsgs = [...messages, { role: 'user', text: inputMsg }];
    setMessages(newMsgs);
    setInputMsg('');

    setTimeout(() => {
      setMessages([
        ...newMsgs,
        { role: 'bot', text: 'DocuDecky ti permette di convertire qualsiasi file in slide ed estratti. Hai bisogno di aiuto con l’abbonamento o con l’upload dei file?' }
      ]);
    }, 600);
  };

  return (
