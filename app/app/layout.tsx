import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "DocuDecky",
  description: "Generatore di documenti e presentazioni con AI",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    
      {children}
    
  );
}
