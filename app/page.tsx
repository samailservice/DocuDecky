'use client';

import React, { useState } from 'react';

export default function Page() {
  const [loading, setLoading] = useState(false);
  const [questions, setQuestions] = useState([]);
  const [answers, setAnswers] = useState('');

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);

    const formElement = e.currentTarget;
    const formData = new FormData(formElement);
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
          return;
        }
      }

      if (!res.ok) {
        const errData = await res.json().catch(() => ({}));
        alert(errData.message || 'Errore durante la generazione');
        setLoading(false);
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
    }
  };

  return ()
