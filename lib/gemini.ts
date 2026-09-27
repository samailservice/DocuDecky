import { GoogleGenAI } from '@google/genai';

const apiKey = process.env.GEMINI_API_KEY || '';

if (!apiKey) {
  console.warn("Attenzione: GEMINI_API_KEY non è configurata nelle variabili d'ambiente.");
}

export const ai = new GoogleGenAI({ apiKey });
