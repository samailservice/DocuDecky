'use client';

import React, { useState } from 'react';
import { Upload, Sparkles, Loader2, ShieldCheck, Zap, Presentation, User, Lock, Mail, ArrowRight, LogOut } from 'lucide-react';

export default function Home() {
  // Stati di Autenticazione
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [authMode, setAuthMode] = useState<'login' | 'register'>('register');
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [name, setName] = useState('');

  // Stati del Tool di Conversione
  const [file, setFile] = useState(null);
  const [sector, setSector] = useState('Business / Aziendale');
  const [objective, setObjective] = useState('Presentazione PPTX');
  const [loading, setLoading] = useState(false);
  const [statusText, setStatusText] = useState('');

  const handleAuth = (e: React.FormEvent) => {
    e.preventDefault();
    // Simulazione di registrazione / login riuscita (gratuita)
    if (!email || !password) {
      alert('Compila tutti i campi obbligatori.');
      return;
    }
    setIsAuthenticated(true);
  };

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

  // Se l'utente NON è autenticato: mostra la Landing Page con spiegazione e form di registrazione/login
  if (!isAuthenticated) {
    return (
      
        {/* Navbar */}
