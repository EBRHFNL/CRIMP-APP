import React, { useState, useMemo, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  Search,
  Star,
  Printer,
  X,
  ChevronRight,
  AlertTriangle,
  Info,
  Upload,
  Download,
  RotateCcw,
  CheckCircle2,
  FileText,
  Link2,
  Pencil,
  Trash2,
  Plus,
  Save,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/*  Voorbeelddataset                                                  */
/*  Gebaseerd op een selectie uit de HANSA-FLEX perslijst - Staal      */
/*  (2026-07). Uitsluitend als voorbeeld voor deze conceptdemo,        */
/*  géén geldige productiespecificatie.                                */
/*  Let op: series/ferrule/standard/mfrCodes zijn technische codes en  */
/*  blijven ongewijzigd, ongeacht de gekozen interfacetaal.            */
/* ------------------------------------------------------------------ */
const DEFAULT_DATA = [
  { id: "te100-06", series: "TE 100", standard: "EN 854 2TE", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["5", "35"], ferrule: "HF-xxx PHT06", crimp: 15.0, tol: null, mandrel: null, d1: 19.0, lf: 28.0, revDate: "2026-07", revSeq: 1 },
  { id: "te100-10", series: "TE 100", standard: "EN 854 2TE", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 10, mfrCodes: ["5", "35", "mw"], ferrule: "HF-xxx PHT10", crimp: 18.8, tol: null, mandrel: null, d1: 23.0, lf: 29.5, revDate: "2026-07", revSeq: 1 },
  { id: "te100-13", series: "TE 100", standard: "EN 854 2TE", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 13, mfrCodes: ["5", "35"], ferrule: "HF-xxx PHT13", crimp: 22.1, tol: null, mandrel: null, d1: 27.0, lf: 31.0, revDate: "2026-07", revSeq: 1 },
  { id: "te100-20", series: "TE 100", standard: "EN 854 2TE", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 20, mfrCodes: ["5", "35"], ferrule: "HF-xxx PHT20", crimp: 30.6, tol: null, mandrel: null, d1: 35.0, lf: 37.5, revDate: "2026-07", revSeq: 1 },

  { id: "te200b-04", series: "TE 200 B - (35)", standard: "EN 854 2TE", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 4, mfrCodes: ["35"], ferrule: "HF-xxx PHT204", crimp: 14.0, tol: null, mandrel: null, d1: 17.0, lf: 27.3, revDate: "2026-07", revSeq: 1 },
  { id: "te200b-13", series: "TE 200 B - (35)", standard: "EN 854 2TE", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 13, mfrCodes: ["35"], ferrule: "HF-xxx PHT13", crimp: 22.5, tol: null, mandrel: null, d1: 27.0, lf: 31.0, revDate: "2026-07", revSeq: 1 },
  { id: "te200b-32", series: "TE 200 B - (35)", standard: "EN 854 2TE", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 32, mfrCodes: ["35"], ferrule: "HF-xxx PHT32", crimp: 44.8, tol: null, mandrel: null, d1: 48.0, lf: 55.0, revDate: "2026-07", revSeq: 1 },

  { id: "hd100ns100-06a", series: "HD 100 - no skive - PHN 100", standard: "EN 853 1SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["25", "27", "35"], ferrule: "HF-xxx PHN106", crimp: 16.6, tol: null, mandrel: "B", d1: 19.6, lf: 30.2, revDate: "2026-07", revSeq: 1 },
  { id: "hd100ns100-06b", series: "HD 100 - no skive - PHN 100", standard: "EN 853 1SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["50"], ferrule: "HF-xxx PHN106", crimp: 16.8, tol: null, mandrel: "B", d1: 19.6, lf: 30.2, revDate: "2026-07", revSeq: 1 },
  { id: "hd100ns100-10a", series: "HD 100 - no skive - PHN 100", standard: "EN 853 1SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 10, mfrCodes: ["25", "27", "35"], ferrule: "HF-xxx PHN110", crimp: 19.8, tol: null, mandrel: "B", d1: 23.2, lf: 32.0, revDate: "2026-07", revSeq: 1 },
  { id: "hd100ns100-10b", series: "HD 100 - no skive - PHN 100", standard: "EN 853 1SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 10, mfrCodes: ["50"], ferrule: "HF-xxx PHN110", crimp: 19.9, tol: null, mandrel: "A", d1: 23.2, lf: 32.0, revDate: "2026-07", revSeq: 1, note: "Bij fabrikant 50 wordt mandrel A gebruikt i.p.v. B. Controleer altijd de actuele officiële perstabel." },
  { id: "hd100ns200-20", series: "HD 100 - no skive - PHN 200", standard: "EN 853 1SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 20, mfrCodes: ["25", "27", "35", "50"], ferrule: "HF-xxx PHN220", crimp: 31.4, tol: "± 0,1 mm", mandrel: null, d1: 37.0, lf: 42.5, revDate: "2026-07", revSeq: 1 },

  { id: "hd100s-06", series: "HD 100 - skive", standard: "EN 853 1SN", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 6, mfrCodes: ["25", "27", "35", "50"], ferrule: "HF-xxx PHD106", crimp: 17.8, tol: null, mandrel: "B", d1: 20, lf: 31.0, revDate: "2026-07", revSeq: 1, skiveLength: 23.0 },
  { id: "hd100s-10a", series: "HD 100 - skive", standard: "EN 853 1SN", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 10, mfrCodes: ["25", "27", "35"], ferrule: "HF-xxx PHD110", crimp: 21.7, tol: null, mandrel: "B", d1: 24, lf: 31.0, revDate: "2026-07", revSeq: 1, skiveLength: 23.0 },
  { id: "hd100s-10b", series: "HD 100 - skive", standard: "EN 853 1SN", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 10, mfrCodes: ["50"], ferrule: "HF-xxx PHD110", crimp: 22.1, tol: null, mandrel: "B", d1: 24, lf: 31.0, revDate: "2026-07", revSeq: 1, skiveLength: 23.0 },

  { id: "kp100ns-06a", series: "KP 100 - no skive", standard: "EN 857 1SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["25", "27", "35"], ferrule: "HF-xxx PHN106", crimp: 16.0, tol: null, mandrel: "B", d1: 19.6, lf: 30.2, revDate: "2026-07", revSeq: 1 },
  { id: "kp100ns-06b", series: "KP 100 - no skive", standard: "EN 857 1SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["50"], ferrule: "HF-xxx PHN106", crimp: 15.7, tol: null, mandrel: "A", d1: 19.6, lf: 30.2, revDate: "2026-07", revSeq: 1 },
  { id: "kp100ns-13a", series: "KP 100 - no skive", standard: "EN 857 1SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 13, mfrCodes: ["25", "35", "50"], ferrule: "HF-xxx PHN113", crimp: 22.8, tol: null, mandrel: "B", d1: 27.5, lf: 34.0, revDate: "2026-07", revSeq: 1 },
  { id: "kp100ns-13b", series: "KP 100 - no skive", standard: "EN 857 1SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 13, mfrCodes: ["27"], ferrule: "HF-xxx PHN113", crimp: 23.5, tol: null, mandrel: "B", d1: 27.5, lf: 34.0, revDate: "2026-07", revSeq: 1 },
  { id: "kp100ns-20", series: "KP 100 - no skive", standard: "EN 857 1SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 20, mfrCodes: ["25", "27", "35"], ferrule: "HF-xxx PHN120", crimp: 30.4, tol: "± 0,1 mm", mandrel: null, d1: 36.0, lf: 42.5, revDate: "2026-07", revSeq: 1 },

  { id: "hd200ns-06", series: "HD 200 - no skive", standard: "EN 853 2SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["25", "27", "35", "50"], ferrule: "HF-xxx PHN206", crimp: 19.0, tol: null, mandrel: "B", d1: 23.0, lf: 30.0, revDate: "2026-07", revSeq: 1 },
  { id: "hd200ns-13a", series: "HD 200 - no skive", standard: "EN 853 2SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 13, mfrCodes: ["25", "27", "35"], ferrule: "HF-xxx PHN213", crimp: 24.6, tol: null, mandrel: "B", d1: 29.0, lf: 32.0, revDate: "2026-07", revSeq: 1 },
  { id: "hd200ns-13b", series: "HD 200 - no skive", standard: "EN 853 2SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 13, mfrCodes: ["50"], ferrule: "HF-xxx PHN213", crimp: 24.5, tol: null, mandrel: "A", d1: 29.0, lf: 32.0, revDate: "2026-07", revSeq: 1 },
  { id: "hd200ns-25a", series: "HD 200 - no skive", standard: "EN 853 2SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 25, mfrCodes: ["25"], ferrule: "HF-xxx PHN225", crimp: 39.2, tol: "± 0,1 mm", mandrel: null, d1: 46.0, lf: 51.0, revDate: "2026-07", revSeq: 1 },
  { id: "hd200ns-25b", series: "HD 200 - no skive", standard: "EN 853 2SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 25, mfrCodes: ["27", "35"], ferrule: "HF-xxx PHN225", crimp: 40.5, tol: null, mandrel: "B", d1: 46.0, lf: 51.0, revDate: "2026-07", revSeq: 1 },
  { id: "hd200ns-25c", series: "HD 200 - no skive", standard: "EN 853 2SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 25, mfrCodes: ["50"], ferrule: "HF-xxx PHN225", crimp: 39.7, tol: null, mandrel: "A", d1: 46.0, lf: 51.0, revDate: "2026-07", revSeq: 1 },

  { id: "kp200ns-06", series: "KP 200 - no skive", standard: "EN 857 2SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["5", "25", "27", "35", "40", "50"], ferrule: "HF-xxx PHN206", crimp: 18.5, tol: null, mandrel: "B", d1: 23.0, lf: 30.0, revDate: "2025-12", revSeq: 1 },
  { id: "kp200ns-08a", series: "KP 200 - no skive", standard: "EN 857 2SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 8, mfrCodes: ["5", "25", "27", "35", "40"], ferrule: "HF-xxx PHN208", crimp: 19.0, tol: null, mandrel: "B", d1: 24.0, lf: 30.0, revDate: "2025-12", revSeq: 1 },
  { id: "kp200ns-08b", series: "KP 200 - no skive", standard: "EN 857 2SC", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 8, mfrCodes: ["50"], ferrule: "HF-xxx PHN208", crimp: 18.9, tol: null, mandrel: "B", d1: 24.0, lf: 30.0, revDate: "2025-12", revSeq: 1 },

  // NY 7206 (fabrikantcode 37) — interlock huls/insert, voorbeeld op basis van perslijst p.51.
  // Eigen (custom) montage- en controlestappen i.p.v. het generieke stappenplan, zie CUSTOM_STEPS.
  { id: "ny7206-37", series: "NY 7206", standard: "NY 7206 met PHY 7206 & PNY 7206", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 6, mfrCodes: ["37"], ferrule: "PHY 7206 (2 striae)", crimp: 18.3, tol: "± 0,1 mm", mandrel: null, d1: 22.0, lf: 46.0, skiveLength: 34.5, assemblyType: "custom", revDate: "2026-08", revSeq: 1 },

  // HD 700 - HD 700 PRO (SAE 100 R15), fabrikantcode 25 — interlock-huls, montage
  // conform de "One Piece Fitting"-instructie (perslijst p.36 en p.38).
  { id: "hd700-10v", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 10, mfrCodes: ["25"], ferrule: "HF-xxx PHD710 (Typ „V“)", crimp: 22.8, tol: "± 0,2 mm", mandrel: null, d1: 27.0, lf: 41.0, skiveLength: 31.0, skiveLengthInt: 11.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "hd700-13v", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 13, mfrCodes: ["25"], ferrule: "HF-xxx PHD713 (Typ „V“)", crimp: 25.6, tol: "± 0,2 mm", mandrel: null, d1: 29.0, da: 30.0, lf: 47.0, skiveLength: 40.0, skiveLengthInt: 13.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "hd700-16v", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 16, mfrCodes: ["25"], ferrule: "HF-xxx PHD716 (Typ „V“)", crimp: 29.6, tol: "± 0,2 mm", mandrel: null, d1: 34.2, da: 35.0, lf: 55.0, skiveLength: 40.0, skiveLengthInt: 15.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "hd700-20l", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["25"], ferrule: "HF-xxx PHD720 / MPF3 3/4 (Typ „L“)", crimp: 35.3, tol: "± 0,2 mm", mandrel: null, d1: 40.0, lf: 62.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "hd700-20v", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["25"], ferrule: "HF-xxx PHD720 (Typ „V“)", crimp: 33.2, tol: "± 0,2 mm", mandrel: null, d1: 38.8, da: 39.6, lf: 68.0, skiveLength: 48.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "PA700H (Typ „V“) wordt geperst als PA700 (Typ „V“). Controleer altijd de actuele officiële perstabel." },
  { id: "hd700-25l", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["25"], ferrule: "HF-xxx PHD725 / MPF3 1 (Typ „L“)", crimp: 41.6, tol: "± 0,2 mm", mandrel: null, d1: 47.0, lf: 75.0, skiveLength: 57.0, skiveLengthInt: 18.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "hd700-25v", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["25"], ferrule: "HF-xxx PHD725 (Typ „V“)", crimp: 43.5, tol: "± 0,2 mm", mandrel: null, d1: 50.0, lf: 73.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "hd700-32l", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 32, mfrCodes: ["25"], ferrule: "HF-xxx PHD732 / MPF4 1.1/4 (Typ „L“)", crimp: 54.3, tol: "± 0,2 mm", mandrel: null, d1: 61.0, lf: 82.0, skiveLength: 64.0, skiveLengthInt: 21.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "hd700-32v", series: "HD 700 - HD 700 PRO (SAE 100 R15)", standard: "SAE 100 R15", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 32, mfrCodes: ["25"], ferrule: "HF-xxx PHD732 (Typ „V“)", crimp: 54.0, tol: "± 0,2 mm", mandrel: null, d1: 59.0, da: 61.0, lf: 78.0, skiveLength: 64.0, skiveLengthInt: 21.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },

  // KP 610 - KP 625 met PA 700, en KP 632 met PA 532 A (interlock, Typ "L" = Lagra / Typ "V" = Voss).
  // Voorbeeld op basis van perslijst p.31.
  { id: "kp610625-10v", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 10, mfrCodes: ["25"], ferrule: "HF-xxx PHD710 (Typ „V“)", crimp: 22.9, tol: "± 0,2 mm", mandrel: null, d1: 27.0, lf: 41.0, skiveLength: 31.0, skiveLengthInt: 11.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-13v", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 13, mfrCodes: ["25"], ferrule: "HF-xxx PHD713 (Typ „V“)", crimp: 25.9, tol: "± 0,2 mm", mandrel: null, d1: 29.0, da: 30.0, lf: 47.0, skiveLength: 40.0, skiveLengthInt: 13.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-16v", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 16, mfrCodes: ["25", "35"], ferrule: "HF-xxx PHD716 (Typ „V“)", crimp: 29.9, tol: "± 0,2 mm", mandrel: null, d1: 34.2, da: 35.0, lf: 55.0, skiveLength: 40.0, skiveLengthInt: 15.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-20l-25", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["25"], ferrule: "HF-xxx PHD720 / MPF3 3/4 (Typ „L“)", crimp: 34.5, tol: "± 0,2 mm", mandrel: null, d1: 40.0, lf: 62.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-20l-35", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["35"], ferrule: "HF-xxx PHD720 / MPF3 3/4 (Typ „L“)", crimp: 34.4, tol: "± 0,2 mm", mandrel: null, d1: 40.0, lf: 62.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-20v-25", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["25"], ferrule: "HF-xxx PHD720 (Typ „V“)", crimp: 34.1, tol: "± 0,2 mm", mandrel: null, d1: 38.8, da: 39.6, lf: 68.0, skiveLength: 48.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-20v-35", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["35"], ferrule: "HF-xxx PHD720 (Typ „V“)", crimp: 34.0, tol: "± 0,2 mm", mandrel: null, d1: 38.8, da: 39.6, lf: 68.0, skiveLength: 48.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-25l", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["25", "35"], ferrule: "HF-xxx PHD725 / MPF3 1 (Typ „L“)", crimp: 41.8, tol: "± 0,2 mm", mandrel: null, d1: 47.0, lf: 75.0, skiveLength: 57.0, skiveLengthInt: 18.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-25v-25", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["25"], ferrule: "HF-xxx PHD725 (Typ „V“)", crimp: 43.1, tol: "± 0,2 mm", mandrel: null, d1: 50.0, lf: 73.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp610625-25v-35", series: "KP 610 - KP 625 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["35"], ferrule: "HF-xxx PHD725 (Typ „V“)", crimp: 42.8, tol: "± 0,2 mm", mandrel: null, d1: 50.0, lf: 73.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp632-32l", series: "KP 632 (PA 532 A)", standard: "PA 532 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 32, mfrCodes: ["25", "35"], ferrule: "HF-xxx PHD532A / MPF3 1.1/4 (Typ „L“)", crimp: 49.7, tol: "± 0,2 mm", mandrel: null, d1: 56.0, lf: 82.0, skiveLength: 64.0, skiveLengthInt: 21.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp632-32v-25", series: "KP 632 (PA 532 A)", standard: "PA 532 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 32, mfrCodes: ["25"], ferrule: "HF-xxx PHD532A (Typ „V“)", crimp: 50.5, tol: "± 0,2 mm", mandrel: null, d1: 55.6, da: 56.0, lf: 81.0, skiveLength: 64.0, skiveLengthInt: 21.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp632-32v-35", series: "KP 632 (PA 532 A)", standard: "PA 532 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 32, mfrCodes: ["35"], ferrule: "HF-xxx PHD532A (Typ „V“)", crimp: 50.0, tol: "± 0,2 mm", mandrel: null, d1: 55.6, da: 56.0, lf: 81.0, skiveLength: 64.0, skiveLengthInt: 21.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },

  // KP 700 met PA 700 (interlock). Voorbeeld op basis van perslijst p.32.
  { id: "kp700-20l-25", series: "KP 700 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["25"], ferrule: "HF-xxx PHD720 / MPF3 3/4 (Typ „L“)", crimp: 35.0, tol: "± 0,2 mm", mandrel: null, d1: 40.0, lf: 62.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp700-20l-35", series: "KP 700 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["35"], ferrule: "HF-xxx PHD720 / MPF3 3/4 (Typ „L“)", crimp: 34.5, tol: "± 0,2 mm", mandrel: null, d1: 40.0, lf: 62.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp700-20v", series: "KP 700 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["25", "35"], ferrule: "HF-xxx PHD720 (Typ „V“)", crimp: 33.7, tol: "± 0,2 mm", mandrel: null, d1: 38.8, da: 39.6, lf: 68.0, skiveLength: 48.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp700-25l", series: "KP 700 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["25"], ferrule: "HF-xxx PHD725 / MPF3 1 (Typ „L“)", crimp: 41.5, tol: "± 0,2 mm", mandrel: null, d1: 47.0, lf: 75.0, skiveLength: 57.0, skiveLengthInt: 18.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp700-25v-25", series: "KP 700 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["25"], ferrule: "HF-xxx PHD725 (Typ „V“)", crimp: 43.4, tol: "± 0,2 mm", mandrel: null, d1: 50.0, lf: 73.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },
  { id: "kp700-25v-35", series: "KP 700 (PA 700)", standard: "PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["35"], ferrule: "HF-xxx PHD725 (Typ „V“)", crimp: 43.0, tol: "± 0,2 mm", mandrel: null, d1: 50.0, lf: 73.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },

  // HD 650 met PA 650 (interlock). Voorbeeld op basis van perslijst p.32.
  { id: "hd650-50l", series: "HD 650 (PA 650)", standard: "EN 856 R13 / PA 650", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 50, mfrCodes: ["5"], ferrule: "HF-xxx PHD650 / MPF4 2 (Typ „L“)", crimp: 77.8, tol: "± 0,2 mm", mandrel: null, d1: 85.0, lf: 100.0, skiveLength: 82.0, skiveLengthInt: 26.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1 },

  // HD 520 - HD 525 met PA 700, en HD 532 - HD 550 met PA 500 A (interlock).
  // Voorbeeld op basis van perslijst p.33. Let op: voor fabrikant (25) gelden
  // afwijkende persmaten (zie HD 500 fabrikant 25) en zijn hier niet opgenomen.
  { id: "hd520525-20l", series: "HD 520 - HD 525 (PA 700)", standard: "EN 856 4SH / PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD720 / MPF3 3/4 (Typ „L“)", crimp: 35.4, tol: "± 0,2 mm", mandrel: null, d1: 40.0, lf: 62.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd520525-20v", series: "HD 520 - HD 525 (PA 700)", standard: "EN 856 4SH / PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 20, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD720 (Typ „V“)", crimp: 33.9, tol: "± 0,2 mm", mandrel: null, d1: 38.8, da: 39.6, lf: 68.0, skiveLength: 48.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd520525-25l", series: "HD 520 - HD 525 (PA 700)", standard: "EN 856 4SH / PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD725 / MPF3 1 (Typ „L“)", crimp: 41.9, tol: "± 0,2 mm", mandrel: null, d1: 47.0, lf: 75.0, skiveLength: 57.0, skiveLengthInt: 18.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd520525-25v", series: "HD 520 - HD 525 (PA 700)", standard: "EN 856 4SH / PA 700", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 25, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD725 (Typ „V“)", crimp: 43.1, tol: "± 0,2 mm", mandrel: null, d1: 50.0, lf: 73.0, skiveLength: 52.0, skiveLengthInt: 17.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd532550-32l", series: "HD 532 - HD 550 (PA 500 A)", standard: "EN 856 4SH / PA 500 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 32, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD532A / MPF3 1.1/4 (Typ „L“)", crimp: 49.8, tol: "± 0,2 mm", mandrel: null, d1: 56.0, lf: 82.0, skiveLength: 64.0, skiveLengthInt: 21.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd532550-32v", series: "HD 532 - HD 550 (PA 500 A)", standard: "EN 856 4SH / PA 500 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 32, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD532A (Typ „V“)", crimp: 50.0, tol: "± 0,2 mm", mandrel: null, d1: 55.6, da: 56.0, lf: 81.0, skiveLength: 64.0, skiveLengthInt: 21.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd532550-40l", series: "HD 532 - HD 550 (PA 500 A)", standard: "EN 856 4SH / PA 500 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 40, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD540A / MPF3 1.1/2 (Typ „L“)", crimp: 56.8, tol: "± 0,2 mm", mandrel: null, d1: 63.0, lf: 93.0, skiveLength: 79.0, skiveLengthInt: 25.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd532550-40v", series: "HD 532 - HD 550 (PA 500 A)", standard: "EN 856 4SH / PA 500 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 40, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD540A (Typ „V“)", crimp: 56.3, tol: "± 0,2 mm", mandrel: null, d1: 62.4, da: 63.0, lf: 96.0, skiveLength: 69.0, skiveLengthInt: 27.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd532550-50l", series: "HD 532 - HD 550 (PA 500 A)", standard: "EN 856 4SH / PA 500 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 50, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD550A / MPF3 2 (Typ „L“)", crimp: 73.8, tol: "± 0,2 mm", mandrel: null, d1: 80.0, lf: 100.0, skiveLength: 82.0, skiveLengthInt: 26.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd532550-50v", series: "HD 532 - HD 550 (PA 500 A)", standard: "EN 856 4SH / PA 500 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 50, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD550A (Typ „V“)", crimp: 72.0, tol: "± 0,2 mm", mandrel: null, d1: 77.4, da: 78.0, lf: 98.5, skiveLength: 72.0, skiveLengthInt: 29.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },
  { id: "hd532550-50v-4rillen", series: "HD 532 - HD 550 (PA 500 A)", standard: "EN 856 4SH / PA 500 A", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 50, mfrCodes: ["5", "10", "34", "mw"], ferrule: "HF-xxx PHD550A (Typ „V“) 4 groeven", crimp: 72.7, tol: "± 0,2 mm", mandrel: null, d1: 77.4, da: 78.0, lf: 98.5, skiveLength: 85.0, skiveLengthInt: 29.0, assemblyType: "interlock", revDate: "2026-08", revSeq: 1, note: "4-groevenvariant. Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel." },

  // RVS (VA) voorbeelddata, gebaseerd op een selectie uit de HANSA-FLEX crimp chart - VA
  // (05.2022). Zelfde fittingfamilie als de staalvariant HD 100, maar met eigen matrijzen
  // (...VA) en het materiaalveld "rvs", zodat de RVS-materiaalregels (langere perstijd,
  // zie tabblad "RVS-fittingen") automatisch getoond worden bij deze combinaties.
  { id: "hd100va-ns-06", series: "HD 100 - no skive VA", standard: "EN 853 1SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 6, mfrCodes: ["5", "25", "27", "35"], ferrule: "HF-xxx PHD106VA", crimp: 16.7, tol: "± 0,1 mm", mandrel: null, d1: 20.0, lf: 34.5, material: "rvs", revDate: "2026-08", revSeq: 1 },
  { id: "hd100va-ns-10", series: "HD 100 - no skive VA", standard: "EN 853 1SN", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 10, mfrCodes: ["5", "25", "27", "35"], ferrule: "HF-xxx PHD110VA", crimp: 21.8, tol: "± 0,1 mm", mandrel: null, d1: 25.0, lf: 35.0, material: "rvs", revDate: "2026-08", revSeq: 1 },
  { id: "hd100va-s-06", series: "HD 100 - skive VA", standard: "EN 853 1SN", uitvoering: "geskived", uitvoeringLabel: "geskived", dn: 6, mfrCodes: ["5", "25", "27", "35"], ferrule: "HF-xxx PHD106VA", crimp: 16.4, tol: "± 0,1 mm", mandrel: null, d1: 20.0, lf: 34.5, skiveLength: 25.0, material: "rvs", revDate: "2026-08", revSeq: 1 },

  // Trapsgewijs persen (stage crimping) - voorbeeld op basis van SGB 100 (perslijst p.45).
  // assemblyType "stage": crimp/crimp2/crimp3 zijn de opeenvolgende persdiameters; de
  // laatste (kleinste) diameter geeft de optimale verbinding, zie tabblad "Trapsgewijs persen".
  { id: "sgb100-stage-20", series: "SGB 100 - stage crimping", standard: "exceeds SAE100R4", uitvoering: "standaard", uitvoeringLabel: "geen skive", dn: 20, mfrCodes: ["25"], ferrule: "HF-xxx PHN220", crimp: 33.0, crimp2: 32.5, crimp3: 32.0, tol: null, mandrel: null, d1: 37.0, lf: 42.5, assemblyType: "stage", revDate: "2026-08", revSeq: 1 },
];

// Vaste, eigen montage-/controlestappen voor combinaties met assemblyType "custom"
// (bijv. NY 7206). Per taal, los van de generieke stappen-template in getSteps().
const CUSTOM_STEPS = {
  "ny7206-37": {
    nl: [
      "Omwikkel de zaagplaats met tape (bijv. Scotch nr. 8981, twee lagen, strak aangetrokken) voordat je de slang doorzaagt, zodat de draadlagen niet uitrafelen. Een cirkelzaag met glad profiel is een alternatief.",
      "Zaag de slang haaks op lengte af. De zaagsnede moet in het midden van de ca. 30 mm brede tape liggen.",
      "Reinig en ontbraam beide uiteinden.",
      "Markeer de insteekdiepte van de huls op beide slanguiteinden.",
      "Verwijder de tape vlak voordat je de huls monteert. Schuif huls PHY7206 op de slang tot aan de markering.",
      "Duw insert PNY7206 in de slang tot deze niet verder kan. Controleer of huls en nippel goed in elkaar grijpen: de referentiegroef op het insert mag na correcte montage niet meer zichtbaar zijn.",
      "Let op dat de huls niet verschuift: het hulseinde moet gelijk blijven lopen met de slangmarkering.",
      "Pers de fitting op de nominale crimpdiameter (aanbevolen: 8-matrijzenkop, matrijs 17).",
      "Controleer de crimpdiameter.",
      "Controleer na het persen dat de slangmarkering niet meer zichtbaar is of net samenvalt met het hulseinde. Bij een afstand groter dan 3 mm mag de slangleiding niet gebruikt worden. Voer op elke productiepartij een druktest uit: bij de eerste twee assemblages en daarna bij elke tiende (proefdruk 1440 bar, houdtijd 1 minuut).",
    ],
    en: [
      "Wrap the cut area with tape (e.g. Scotch no. 8981, two wraps, high tension) before cutting the hose, to keep the wires from fraying. A circular saw with a smooth profile is an alternative.",
      "Cut the hose square to length. The cut should be in the middle of the roughly 30 mm wide tape.",
      "Clean and deburr both ends.",
      "Mark the ferrule insert depth on both hose ends.",
      "Remove the tape just before fitting the ferrule. Slide ferrule PHY7206 onto the hose up to the marking.",
      "Push insert PNY7206 into the hose until it stops. Check that ferrule and nipple engage correctly: the reference groove on the insert must no longer be visible once correctly seated.",
      "Make sure the ferrule doesn't shift: the ferrule end must still line up with the hose marking.",
      "Crimp the fitting to the nominal crimp diameter (recommended: 8-die head, die 17).",
      "Check the crimp diameter.",
      "After crimping, check that the hose marking is no longer visible or just coincides with the ferrule end. If the gap exceeds 3 mm, the hose assembly must not be used. Pressure-test each production lot: the first two assemblies, then every tenth one (test pressure 1440 bar, hold time 1 minute).",
    ],
    de: [
      "Die Schnittstelle vor dem Ablängen mit Klebeband umwickeln (z. B. Scotch-Nr. 8981, zwei Lagen, straff gezogen), damit die Drahtlagen nicht ausfransen. Alternativ ein Kreissägeblatt mit glattem Profil verwenden.",
      "Den Schlauch rechtwinklig auf Länge schneiden. Die Schnittfläche sollte mittig im ca. 30 mm breiten Klebeband liegen.",
      "Beide Enden reinigen und entgraten.",
      "Die Einstecktiefe der Hülse auf beiden Schlauchenden markieren.",
      "Das Klebeband erst unmittelbar vor der Hülsenmontage entfernen. Hülse PHY7206 bis zur Markierung auf den Schlauch schieben.",
      "Einsatz PNY7206 bis zum Anschlag in den Schlauch drücken. Prüfen, ob Hülse und Nippel korrekt ineinandergreifen: die Referenznut am Einsatz darf nach korrektem Sitz nicht mehr sichtbar sein.",
      "Darauf achten, dass die Hülse nicht verrutscht: das Hülsenende muss weiterhin mit der Schlauchmarkierung fluchten.",
      "Die Armatur auf den nominalen Crimp-Durchmesser verpressen (empfohlen: 8-Backen-Kopf, Backe 17).",
      "Den Crimp-Durchmesser kontrollieren.",
      "Nach dem Verpressen prüfen, dass die Schlauchmarkierung nicht mehr sichtbar ist bzw. gerade mit dem Hülsenende zusammenfällt. Bei einem Abstand über 3 mm darf die Schlauchleitung nicht verwendet werden. Bei jeder Fertigungscharge eine Druckprüfung durchführen: die ersten zwei Baugruppen, danach jede zehnte (Prüfdruck 1440 bar, Haltezeit 1 Minute).",
    ],
  },
};

// Eén enkele combinatie in de voorbeelddata heeft een bijzondere opmerking.
// Vertaald per taal; buiten de dataset gehouden zodat DEFAULT_DATA neutraal blijft.
const NOTE_TRANSLATIONS = {
  "hd100ns100-10b": {
    nl: "Bij fabrikant 50 wordt mandrel A gebruikt i.p.v. B. Controleer altijd de actuele officiële perstabel.",
    en: "For manufacturer 50, mandrel A is used instead of B. Always check the current official crimping table.",
    de: "Bei Hersteller 50 wird Dorn A statt B verwendet. Bitte immer die aktuelle offizielle Presstabelle prüfen.",
  },
  "hd700-20v": {
    nl: "PA700H (Typ „V“) wordt geperst als PA700 (Typ „V“). Controleer altijd de actuele officiële perstabel.",
    en: "PA700H (Type „V“) is crimped like PA700 (Type „V“). Always check the current official crimping table.",
    de: "PA700H (Typ „V“) wird wie PA700 (Typ „V“) verpresst. Bitte immer die aktuelle offizielle Presstabelle prüfen.",
  },
  "hd520525-20l": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd520525-20v": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd520525-25l": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd520525-25v": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd532550-32l": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd532550-32v": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd532550-40l": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd532550-40v": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd532550-50l": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd532550-50v": { nl: "Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
  "hd532550-50v-4rillen": { nl: "4-groevenvariant. Voor fabrikant 25 gelden afwijkende persmaten (zie HD 500 fabrikant 25). Controleer altijd de actuele officiële perstabel.", en: "4-groove variant. Different crimping dimensions apply for manufacturer 25 (see HD 500 manufacturer 25). Always check the current official crimping table.", de: "4-Rillen-Variante. Für Hersteller 25 gelten abweichende Pressmaße (siehe HD 500 Hersteller 25). Bitte immer die aktuelle offizielle Presstabelle prüfen." },
};

/* ------------------------------------------------------------------ */
/*  Talen                                                             */
/* ------------------------------------------------------------------ */
const LANGUAGES = [
  { code: "nl", flag: "🇳🇱", label: "Nederlands" },
  { code: "en", flag: "🇬🇧", label: "English" },
  { code: "de", flag: "🇩🇪", label: "Deutsch" },
];
const LOCALE_MAP = { nl: "nl-NL", en: "en-GB", de: "de-DE" };

const I18N = {
  nl: {
    // Header / algemeen
    appTitle: "Crimp Data",
    appSubtitle: "Digitale perstabel — concept",
    disclaimerStrong: "Conceptdemo, voorbeelddata.",
    disclaimerRest: "Geen geldige productiespecificaties. Raadpleeg altijd de officiële HANSA-FLEX crimptabellen.",
    btnUpdateData: "Data bijwerken",
    btnPrintExport: "Print / exporteer",
    statusShared: (date, count) => `Gedeelde data actief — bijgewerkt op ${date} (${count} regels)`,
    statusDefault: "Ingebouwde voorbeelddata — nog geen upload actief",
    tabSearch: "Zoeken",
    tabFavorites: "Favorieten",
    tabMandrel: "Test-Doorn",
    tabOnePiece: "One Piece Fitting",
    tabInterlock: "Interlock Fitting",
    tabStageCrimp: "Trapsgewijs persen",
    tabStainless: "RVS-fittingen",
    languageLabel: "Taal",

    // Filters
    filterSeriesLabel: "Slangtype / serie",
    filterSeriesPlaceholder: "Bijv. HD 100, KP 200, TE 100 …",
    filterMfrLabel: "Slang code",
    filterAll: "Alle",
    filterDnLabel: "Nominale diameter (DN)",
    filterVariantLabel: "Uitvoering",
    variantStandardChip: "Standaard (geen skive)",
    variantSkivedChip: "Geskived",
    variantStandardShort: "geen skive",
    variantSkivedShort: "geskived",
    resultsCount: (n) => `${n} resultaten gevonden`,
    btnClearFilters: "Filters wissen",

    // Leegstaten
    emptyResultsTitle: "Geen combinaties gevonden",
    emptyResultsMessage: "Pas de filters aan of wis ze om alle voorbeelddata te tonen.",
    emptyFavoritesTitle: "Nog geen favorieten",
    emptyFavoritesMessage: "Tik op het sterretje bij een resultaat om deze hier snel terug te vinden.",
    favoritesCount: (n) => `${n} favoriete combinatie(s)`,

    // Resultaatkaart
    labelHose: "Huls:",
    labelCrimpO: "Crimp Ø",
    fixedSizeViaDie: "vaste maat via matrijs",
    toleranceLabel: (tol) => `tolerantie ${tol}`,
    labelMandrel: "Mandrel",
    mfrChipPrefix: "Fabr.",
    fittingTypeV: "Type V (Voss)",
    fittingTypeL: "Type L (Lagra)",
    materialChipRvs: "RVS",
    stageChipLabel: "Trapsgewijs",
    revShort: (revDate, seq) => `Rev. ${revDate}-R${String(seq || 1).padStart(2, "0")}`,
    btnDetails: "Details & persinstructie",
    favAriaSave: "Bewaar als favoriet",
    favAriaRemove: "Verwijder uit favorieten",
    noteIndicatorAria: "Bevat een opmerking, zie details",

    // Detailweergave
    fieldNominalDiameter: "Nominale diameter",
    fieldFerruleMarking: "Hulsmarkering",
    fieldMandrel: "Mandrel",
    fieldD1LF: "D1 / LF",
    fieldDA: "DA (buitendiameter huls)",
    fieldMfrCodes: "Fabrikantcode(s)",
    notApplicable: "n.v.t.",
    headingInstructions: "Persinstructie",
    referenceLine: (revDate, seq) => `Referentie: revisie Rev. ${revDate}-R${String(seq || 1).padStart(2, "0")}. Bron: interne perslijst (voorbeelddata).`,
    btnSaveFavorite: "Bewaar als favoriet",
    btnRemoveFavorite: "Verwijder favoriet",
    btnPrintCard: "Print kaart",
    ariaClose: "Sluiten",

    // Persinstructiestappen
    stepMandrel: (mandrel, ferrule, dn) => `Selecteer mandrel/dorn type ${mandrel}, passend bij huls ${ferrule} (DN${dn}).`,
    stepDieTable: (ferrule, dn) => `Kies de matrijs conform de matrijstabel voor huls ${ferrule} (DN${dn}).`,
    stepSkive: (len, intLen) => intLen ? `Skive de slang extern tot ${len} mm en intern tot ${intLen} mm.` : `Skive de slang (extern, en indien van toepassing intern) tot een lengte van ${len} mm.`,
    stepSlideFerrule: "Schuif de huls volledig op de slang tot aan de aanslag en markeer of controleer de positie ten opzichte van de eindstreep.",
    stepSetDiameter: (crimp, tol) =>
      `Stel de crimp-/persdiameter in op ${crimp} mm${tol ? ` (tolerantie ${tol})` : " (vaste maat via matrijs — raadpleeg de officiële tabel voor de exacte tolerantie)"}.`,
    stepSetDiameterStage: (diameters) =>
      `Pers stapsgewijs: bereik eerst Ø1 = ${diameters[0]} mm${diameters.length > 2 ? `, verklein vervolgens naar Ø2 = ${diameters[1]} mm en tot slot naar Ø3 = ${diameters[2]} mm` : `, verklein vervolgens naar Ø2 = ${diameters[1]} mm`}. De laatste (kleinste) diameter geeft de optimale verbinding, zie ook het tabblad "Trapsgewijs persen".`,
    stepPress: "Pers de fitting in één beweging conform de ingestelde diameter.",
    stepVerify: "Verifieer de persmaat met minimaal 4 metingen verdeeld over de omtrek (bijvoorbeeld op 0°, 90°, 180° en 45°/135°).",
    stepVisualCheck: "Controleer visueel op knikken of scheve montage; raadpleeg bij twijfel de officiële tabel voor het uitvoeren van een druktest.",

    // Test-Doorn (mandrel) - naslagpagina, gebaseerd op HANSA-FLEX werkinstructie PN-DORN…
    mandrelPageTitle: "Test-Doorn",
    mandrelPageWarning: "Gebruik de test-doorn uitsluitend als dit in de perstabel is aangegeven.",
    mandrelPageIntro: "De HF test-doorns zijn ontworpen om de persdiameter te controleren bij mogelijke tolerantieafwijkingen van de slang. In de HF-perstabel staat aangegeven of een boordoorsnede-inklapping (bore collapse) gecontroleerd kan worden, en met welke diameter.",
    mandrelFigures: [
      { title: "Figuur 1: correcte montage Ø A", text: "Er is een boordoorsnede-inklapping bereikt. De diameter A van de test-doorn kan tot aan de samendrukking worden ingebracht, maar niet verder dan de zogenaamde \"stap 1\". De persing is correct." },
      { title: "Figuur 2: correcte montage Ø B", text: "Er is een boordoorsnede-inklapping bereikt. De diameter B van de test-doorn kan tot aan de samendrukking worden ingebracht, maar niet verder dan de zogenaamde \"stap 2\". De persing is correct." },
      { title: "Figuur 3: onvoldoende inklapping", text: "De test-doorn gaat volledig in de fitting. De boordoorsnede-inklapping is onvoldoende. De montage is niet correct. Werkwijze: verklein de persdiameter in stappen van 0,1 mm tot een voldoende inklapping is bereikt (zie figuur 1 of 2)." },
      { title: "Figuur 4: overmatige inklapping", text: "De boordoorsnede is overmatig ingeklapt. De extreme vernauwing van de binnendiameter verhindert het inbrengen van de doorn. De montage is niet correct; de slangassemblage mag niet gebruikt worden. Werkwijze: controleer de afmetingen van slang en fitting en pas de persdiameter aan tot een correcte montage (zie figuur 1 of 2)." },
    ],
    mandrelFrequencyHeading: "Testfrequentie",
    mandrelFrequencyText: "Bij unitproductie wordt elke slangassemblage gecontroleerd. Bij seriematige productie worden de eerste 6 assemblages gecontroleerd; daarna volgen steekproeven, minimaal elke 20e assemblage en altijd bij een chargewissel van de slang.",
    mandrelControlHeading: "Controle van de test-doorn",
    mandrelControlText: "Controleer de test-doorn visueel voor elk gebruik. Verbogen of gebroken test-doorns moeten vervangen worden. Een lichte slijtage van diameter A of B is niet kritisch. Bij zware beschadiging moet de doorn direct vervangen worden.",

    // One Piece Fitting - naslagpagina, gebaseerd op HANSA-FLEX montage-instructie PA 700 PLUS / HD 700 PLUS / HD 700 LL
    onePiecePageTitle: "One Piece Fitting — montage-instructie",
    onePiecePageIntro: "Montage-instructie voor interlock-fittingen (one piece fitting), zoals PA 700 PLUS, HD 700 PLUS en HD 700 LL. Van toepassing op alle combinaties met dit fittingtype, zie ook de crimpkaart van de betreffende slang.",
    onePieceSteps: [
      "Neem na het (intern en extern) skiven de insteekdiepte van de slang over uit de perstabel en markeer deze op de slangmantel. Deze markering helpt om te controleren of de slang volledig tot de eindpositie op de fitting/insert is geschoven.",
      "Schuif de huls op de slang.",
      "Let op: bij fittingen/inserts in NW 32, 40 en 50 moeten de 2 O-ringen op de daarvoor bestemde moeren aan het slanguiteinde geplaatst worden. Controleer de juiste montage visueel voordat je de fitting/insert monteert. Fittingen/inserts zonder deze 2 O-ringen mogen niet gemonteerd worden. Breng de fitting/insert aan (handmatig of met een insteekmachine); de binnenbuis van de slang mag niet beschadigd raken. Smeermiddel OELPAG46 kan gebruikt worden om mechanische slijtage te voorkomen; droog monteren is niet toegestaan. Let op de markering van de insteekdiepte op de slangmantel.",
      "Schuif de huls naar voren tot aan de insteekmoer.",
      "Pers zoals aanbevolen.",
      "Controleer na het persen: de montage visueel, de crimpdiameter, en de markering van de insteekdiepte op de slangmantel (een kleine speling door materiaalvervorming is toegestaan).",
    ],
    onePieceWarningText: "Een te grote opening tussen huls en slangmarkering wijst op een verkeerde montage: het interlockgedeelte van de huls grijpt dan niet goed in de spiraalwapening.",
    onePieceNoteText: "Het gebruik van een voormontagemachine bevordert een correcte montage. Een veelvoorkomende storingsoorzaak bij interlockfittingen is een te lange interne skivelengte.",
    onePieceLinkText: "Deze combinatie wordt gemonteerd volgens de One Piece Fitting-instructie.",
    btnViewOnePiece: "Bekijk montage-instructie",

    // Interlock Fitting - naslagpagina, gebaseerd op HANSA-FLEX montage-instructie interlock fittingen (perslijst p.30)
    interlockPageTitle: "Interlock Fitting — montage-instructie",
    interlockPageIntro: "Volg deze stappen zorgvuldig: een verkeerde montage blijft de meest voorkomende oorzaak van storingen bij interlock-fittingen.",
    interlockSteps: [
      "Na het skiven (intern en extern): schuif de huls op de slang.",
      "Controleer de juiste positie van de slang ten opzichte van de huls-kraag (collar).",
      "Markeer de correcte eindpositie van de huls op de slangmantel.",
      "Steek vóór het persen de fitting/insert in tot aan de huls-kraag.",
      "Controleer de positie van de huls, dat wil zeggen de overeenstemming met de markering op de slangmantel.",
      "Pers zoals aanbevolen.",
      "Controleer na het persen de hulspositie ten opzichte van de slangmarkering; een kleine speling door materiaalvervorming is toegestaan.",
    ],
    interlockWarningText: "Een te grote opening tussen huls en slangmarkering wijst op een verkeerde montage: het interlockgedeelte van de huls grijpt dan niet goed in de spiraalwapening.",
    interlockNoteText: "Het gebruik van een voormontagemachine bevordert een correcte montage. Een veelvoorkomende storingsoorzaak bij interlockfittingen is een te lange interne skivelengte.",

    // Trapsgewijs persen - naslagpagina, gebaseerd op HANSA-FLEX perslijst (staal p.45, RVS p.21):
    // combinaties met assemblyType "stage" moeten in twee of drie stappen naar de eindmaat geperst worden.
    stageCrimpPageTitle: "Trapsgewijs persen",
    stageCrimpPageIntro: "Bij een aantal slang/fitting-combinaties (bijvoorbeeld SGB 100 en NY 1800, in zowel staal als RVS) mag de eindmaat niet in één keer geperst worden. Doorloop de crimpdiameters in de aangegeven volgorde; bij het bereiken van de laatste (kleinste) diameter is de verbinding optimaal.",
    stageCrimpWhyHeading: "Waarom stapsgewijs?",
    stageCrimpWhyText: "Bij deze combinaties is de sprong tussen de ongeperste en de uiteindelijke diameter te groot om in één beweging te persen zonder de wapening van de slang te beschadigen. Door via een of twee tussenmaten te persen wordt de vervorming geleidelijk opgebouwd.",
    stageCrimpExampleHeading: "Voorbeeld: SGB 100, DN 20",
    stageCrimpExampleText: "Ø1 = eerste (grootste) persdiameter, Ø2 = tussenmaat, Ø3 = eindmaat. Zie de kaart van deze combinatie voor de exacte waarden.",
    stageCrimpApplicableHeading: "Van toepassing op",
    stageCrimpApplicableText: "SGB 100 (staal en RVS) en NY 1800. Combinaties met dit persschema zijn herkenbaar aan het label \"Trapsgewijs\" op de kaart.",
    stageCrimpLinkText: "Deze combinatie wordt stapsgewijs geperst.",
    btnViewStageCrimp: "Bekijk trapsgewijze instructie",

    // RVS-fittingen - naslagpagina, gebaseerd op HANSA-FLEX crimp chart VA (05.2022).
    // Materiaalregels die gelden zodra een combinatie material "rvs" heeft.
    stainlessPageTitle: "RVS-fittingen — materiaalregels",
    stainlessPageIntro: "Voor RVS (VA) slang/fitting-combinaties gelden, naast de gewone persinstructie, twee extra materiaalregels. Combinaties met deze materiaalregels zijn herkenbaar aan het label \"RVS\" op de kaart.",
    stainlessHoldTimeHeading: "Langere houdtijd in de pers",
    stainlessHoldTimeText: "De houdtijd van RVS-fittingen in de pers moet na het bereiken van de crimpdiameter langer zijn dan bij staal (circa 5 seconden), om voldoende materiaalvervorming te garanderen.",
    stainlessDisclaimerHeading: "Persmaten zijn standaardwaarden",
    stainlessDisclaimerText: "Persdiameters zijn uitsluitend standaardwaarden. De boordoorsnede-inklapping (bore collapse) moet altijd gecontroleerd worden, bijvoorbeeld met de test-doorn; zie ook de interne Werknorm hoofdstuk 8.",
    stainlessInterlockDiffHeading: "Afwijking bij interlock-montage",
    stainlessInterlockDiffText: "Bij RVS interlock-fittingen moet, na het opschuiven van de huls tot de eerste weerstand (de \"bump\") van de interlock, deze weerstand bewust overwonnen worden. Trek de huls daarna weer iets naar voren zodat de opening met de fitting sluit, vóórdat je gaat persen.",
    stainlessAutoNoteText: "Let op: dit is een RVS-fitting. De houdtijd in de pers moet langer zijn (ca. 5 sec.).",
    stainlessLinkText: "Voor deze combinatie gelden de RVS-materiaalregels.",
    btnViewStainlessInfo: "Bekijk RVS-materiaalregels",

    // Wijzigen (bewerken/toevoegen/verwijderen van combinaties)
    btnEdit: "Wijzigen",
    btnAddCombination: "Nieuwe combinatie",
    editTitleNew: "Nieuwe combinatie toevoegen",
    editTitleEdit: "Combinatie wijzigen",
    editFieldSeries: "Slangtype / serie",
    editFieldStandard: "Norm",
    editFieldUitvoering: "Uitvoering",
    editFieldDn: "Nominale diameter (DN)",
    editFieldMfrCodes: "Fabrikantcode(s)",
    editFieldMfrAdd: "Code toevoegen en Enter",
    editFieldFerrule: "Huls / matrijs",
    editFieldCrimp: "Crimpdiameter (mm)",
    editFieldTol: "Tolerantie (optioneel)",
    editFieldMandrel: "Mandrel (optioneel)",
    editFieldD1: "D1 (mm)",
    editFieldDa: "DA (mm, optioneel)",
    editFieldLf: "LF (mm)",
    editFieldSkive: "Skive-lengte extern (mm, optioneel)",
    editFieldSkiveInt: "Skive-lengte intern (mm, optioneel)",
    editFieldAssemblyType: "Montagetype",
    assemblyTypeStandard: "Standaard",
    assemblyTypeInterlock: "One Piece Fitting (interlock)",
    assemblyTypeStage: "Trapsgewijs persen",
    editFieldMaterial: "Materiaal",
    materialLabelStaal: "Staal",
    materialLabelRvs: "RVS (VA)",
    editFieldCrimp2: "Crimpdiameter Ø2 (mm, bij trapsgewijs persen)",
    editFieldCrimp3: "Crimpdiameter Ø3 (mm, optioneel)",
    editFieldNote: "Opmerking (optioneel)",
    btnSaveEdit: "Opslaan",
    btnCancelEdit: "Annuleren",
    btnDeleteItem: "Combinatie verwijderen",
    btnDeleteConfirm: "Klik nogmaals om definitief te verwijderen",
    editErrorRequired: "Vul alle verplichte velden in (series, norm, DN, fabrikantcode, huls, crimpdiameter, D1, LF).",
    editSavedRev: (revDate, seq) => `Opgeslagen als Rev. ${revDate}-R${String(seq).padStart(2, "0")}.`,

    // Data bijwerken
    updateTitle: "Data bijwerken",
    storageWarning: "Gedeelde opslag is in deze omgeving niet beschikbaar. Een upload werkt nu alleen tijdelijk, voor dit scherm.",
    introShared: "Upload een .xlsx-bestand met de bijgewerkte crimpdata (gebruik het sjabloon hieronder als basis). Na controle en bevestiging wordt de data gedeeld opgeslagen — iedereen die deze tool opent ziet daarna direct de nieuwe dataset. Revisienummers worden automatisch bijgewerkt.",
    introLocal: "Upload een .xlsx-bestand met de bijgewerkte crimpdata (gebruik het sjabloon hieronder als basis). Na controle en bevestiging wordt de data tijdelijk in dit scherm gebruikt. Revisienummers worden automatisch bijgewerkt.",
    btnChooseFile: "Bestand kiezen (.xlsx)",
    btnDownloadTemplate: "Sjabloon downloaden",
    btnDownloadData: "Huidige database downloaden",
    errFileType: "Alleen .xlsx-bestanden worden ondersteund.",
    btnAnalyse: "Analyseren & controleren",
    errorsFoundTitle: (n) => `${n} probleem/problemen gevonden`,
    moreErrors: (n) => `… en ${n} meer.`,
    validatedCountSuffix: "combinatie(s) gevalideerd, klaar om op te slaan.",
    tableHeaderSeries: "Serie",
    tableHeaderDn: "DN",
    tableHeaderHose: "Huls",
    moreRows: (n) => `… en ${n} regel(s) meer.`,
    btnCancel: "Annuleren",
    btnConfirmSave: "Bevestigen en opslaan",
    savingLabel: "Bezig met opslaan…",
    statusUpdatedOn: (date, count) => `Huidige gedeelde data bijgewerkt op ${date} (${count} regels).`,
    statusNoUpload: "Er is nog geen gedeelde upload actief — de tool gebruikt de ingebouwde voorbeelddata.",
    btnResetDefault: "Terugzetten naar voorbeelddata",
    btnResetConfirm: "Klik nogmaals om te bevestigen",
    fileReadError: "Kon het bestand niet lezen.",
    saveFailedDefault: "Opslaan is mislukt.",

    // Validatieberichten (upload)
    errEmptyInput: "Geen data gevonden — het xlsx-bestand is leeg.",
    errNoDataRows: "Geen datarijen gevonden.",
    errMissingColumns: (list) => `Verplichte kolom(men) ontbreken: ${list}.`,
    errSeriesMissing: (row) => `Rij ${row}: 'series' ontbreekt.`,
    errStandardMissing: (row) => `Rij ${row}: 'standard' ontbreekt.`,
    errUitvoeringInvalid: (row, found) => `Rij ${row}: 'uitvoering' moet 'standaard' of 'geskived' zijn (gevonden: "${found}").`,
    errDnInvalid: (row) => `Rij ${row}: 'dn' is geen geldig getal.`,
    errMfrCodesMissing: (row) => `Rij ${row}: 'mfrCodes' ontbreekt (scheid meerdere codes met een |, bijv. "25|27|35").`,
    errFerruleMissing: (row) => `Rij ${row}: 'ferrule' ontbreekt.`,
    errCrimpInvalid: (row) => `Rij ${row}: 'crimp' is geen geldig getal.`,
    errD1Invalid: (row) => `Rij ${row}: 'd1' is geen geldig getal.`,
    errLfInvalid: (row) => `Rij ${row}: 'lf' is geen geldig getal.`,
    errRevMissing: (row) => `Rij ${row}: 'rev' ontbreekt.`,
    errRevDateMissing: (row) => `Rij ${row}: 'revDate' ontbreekt.`,

    // Printweergave
    printTitle: "HANSA-FLEX Crimp Data — Referentiekaart",
    printMetaLine: (label, date) => `Concept, voorbeelddata — ${label} — gegenereerd op ${date}`,
    printLabelOne: "1 combinatie",
    printLabelFavorites: "favorieten",
    printLabelResults: "zoekresultaten",
    printThSeries: "Serie",
    printThMfrCode: "Fabr.code",
    printThHose: "Huls",
    printThMandrel: "Mandrel",
    printThD1LF: "D1 / LF",
    printThReference: "Referentie",
    printInstructionsHeading: "Persinstructie",
  },

  en: {
    appTitle: "Crimp Data",
    appSubtitle: "Digital crimping table — concept",
    disclaimerStrong: "Concept demo, sample data.",
    disclaimerRest: "Not valid production specifications. Always consult the official HANSA-FLEX crimping tables.",
    btnUpdateData: "Update data",
    btnPrintExport: "Print / export",
    statusShared: (date, count) => `Shared data active — updated on ${date} (${count} rows)`,
    statusDefault: "Built-in sample data — no upload active yet",
    tabSearch: "Search",
    tabFavorites: "Favourites",
    tabMandrel: "Mandrel",
    tabOnePiece: "One Piece Fitting",
    tabInterlock: "Interlock Fitting",
    tabStageCrimp: "Stage Crimping",
    tabStainless: "Stainless Fittings",
    languageLabel: "Language",

    filterSeriesLabel: "Hose type / series",
    filterSeriesPlaceholder: "E.g. HD 100, KP 200, TE 100 …",
    filterMfrLabel: "Hose code",
    filterAll: "All",
    filterDnLabel: "Nominal diameter (DN)",
    filterVariantLabel: "Version",
    variantStandardChip: "Standard (no skive)",
    variantSkivedChip: "Skived",
    variantStandardShort: "no skive",
    variantSkivedShort: "skived",
    resultsCount: (n) => `${n} results found`,
    btnClearFilters: "Clear filters",

    emptyResultsTitle: "No combinations found",
    emptyResultsMessage: "Adjust or clear the filters to show all sample data.",
    emptyFavoritesTitle: "No favourites yet",
    emptyFavoritesMessage: "Tap the star on a result to find it here quickly.",
    favoritesCount: (n) => `${n} favourite combination(s)`,

    labelHose: "Ferrule:",
    labelCrimpO: "Crimp Ø",
    fixedSizeViaDie: "fixed size via die",
    toleranceLabel: (tol) => `tolerance ${tol}`,
    labelMandrel: "Mandrel",
    mfrChipPrefix: "Mfr.",
    fittingTypeV: "Type V (Voss)",
    fittingTypeL: "Type L (Lagra)",
    materialChipRvs: "Stainless",
    stageChipLabel: "Stage crimp",
    revShort: (revDate, seq) => `Rev. ${revDate}-R${String(seq || 1).padStart(2, "0")}`,
    btnDetails: "Details & press instructions",
    favAriaSave: "Save as favourite",
    favAriaRemove: "Remove from favourites",
    noteIndicatorAria: "Contains a note, see details",

    fieldNominalDiameter: "Nominal diameter",
    fieldFerruleMarking: "Ferrule marking",
    fieldMandrel: "Mandrel",
    fieldD1LF: "D1 / LF",
    fieldDA: "DA (ferrule outer diameter)",
    fieldMfrCodes: "Manufacturer code(s)",
    notApplicable: "n/a",
    headingInstructions: "Press instructions",
    referenceLine: (revDate, seq) => `Reference: revision Rev. ${revDate}-R${String(seq || 1).padStart(2, "0")}. Source: internal crimping table (sample data).`,
    btnSaveFavorite: "Save as favourite",
    btnRemoveFavorite: "Remove favourite",
    btnPrintCard: "Print card",
    ariaClose: "Close",

    stepMandrel: (mandrel, ferrule, dn) => `Select mandrel type ${mandrel}, matching ferrule ${ferrule} (DN${dn}).`,
    stepDieTable: (ferrule, dn) => `Choose the die according to the die table for ferrule ${ferrule} (DN${dn}).`,
    stepSkive: (len, intLen) => intLen ? `Skive the hose externally to ${len} mm and internally to ${intLen} mm.` : `Skive the hose (external, and internal where applicable) to a length of ${len} mm.`,
    stepSlideFerrule: "Slide the ferrule fully onto the hose up to the stop, and mark or check its position relative to the end line.",
    stepSetDiameter: (crimp, tol) =>
      `Set the crimp diameter to ${crimp} mm${tol ? ` (tolerance ${tol})` : " (fixed size via die — consult the official table for the exact tolerance)"}.`,
    stepSetDiameterStage: (diameters) =>
      `Crimp in stages: first reach Ø1 = ${diameters[0]} mm${diameters.length > 2 ? `, then reduce to Ø2 = ${diameters[1]} mm, and finally to Ø3 = ${diameters[2]} mm` : `, then reduce to Ø2 = ${diameters[1]} mm`}. The last (smallest) diameter gives the optimal connection, see also the "Stage Crimping" tab.`,
    stepPress: "Crimp the fitting in a single motion according to the set diameter.",
    stepVerify: "Verify the crimped size with at least 4 measurements spread around the circumference (e.g. at 0°, 90°, 180° and 45°/135°).",
    stepVisualCheck: "Visually check for kinks or misalignment; if in doubt, consult the official table for performing a pressure test.",

    // Mandrel - reference page, based on HANSA-FLEX works standard PN-DORN…
    mandrelPageTitle: "Mandrel",
    mandrelPageWarning: "Use the test mandrel only if notified in the swaging chart.",
    mandrelPageIntro: "The HF test mandrels were designed to check the swaging diameter for potential tolerance discrepancies of the hose. The HF swaging chart indicates if a bore collapse can be controlled, and with which diameter.",
    mandrelFigures: [
      { title: "Figure 1: correct assembly Ø A", text: "A bore collapse has been reached. The diameter A of the test mandrel can be inserted up to the compression, but not farther than the so-called \"step 1\". The swaging is correct." },
      { title: "Figure 2: correct assembly Ø B", text: "A bore collapse has been reached. The diameter B of the test mandrel can be inserted up to the compression, but not farther than the so-called \"step 2\". The swaging is correct." },
      { title: "Figure 3: unsatisfactory bore collapse", text: "The test mandrel goes completely into the fitting. The bore collapse is unsatisfactory. The assembly is not correct. Procedure: reduce the crimp diameter in steps of 0.1 mm until a satisfactory collapse has been reached (cf. figure 1 or 2)." },
      { title: "Figure 4: excessive bore collapse", text: "The bore collapsed excessively. The extreme reduction of the inside bore inhibits mandrel insertion. The assembly is not correct; the hose assembly must not be used. Procedure: check the hose and fitting dimensions and adjust the crimp diameter to get a correct assembly (cf. figure 1 or 2)." },
    ],
    mandrelFrequencyHeading: "Frequency of testing",
    mandrelFrequencyText: "By unit production, every hose assembly is controlled. By series production, the first 6 hose assemblies are controlled; consecutively, random inspections occur at least every 20th assembly, and always on any batch change of hose.",
    mandrelControlHeading: "Test mandrel control",
    mandrelControlText: "Visually examine the test mandrel before any use. Twisted or broken test mandrels have to be replaced. A slight deterioration of diameter A or B is uncritical. If the test mandrel is heavily damaged, it has to be replaced immediately.",

    // One Piece Fitting - reference page, based on HANSA-FLEX assembly instruction for PA 700 PLUS / HD 700 PLUS / HD 700 LL
    onePiecePageTitle: "One Piece Fitting — assembly instruction",
    onePiecePageIntro: "Assembly instruction for interlock (one piece) fittings, such as PA 700 PLUS, HD 700 PLUS and HD 700 LL. Applies to all combinations with this fitting type; see also the crimping card of the relevant hose.",
    onePieceSteps: [
      "After (internal and external) skiving, take the insert depth of the hose from the swaging chart and mark it on the hose cover. This mark helps ensure the hose is fully inserted to the final position on the fitting/insert.",
      "Slip on the ferrule.",
      "Look out: for fittings/inserts in NW 32, 40 and 50, put the 2 O-rings on the specific nuts at the tail end. Check the correct mounting visually before assembling the fitting/insert. Fittings/inserts without these 2 O-rings must not be assembled. Insert the fitting/insert (manually or with an insert machine); the inside tube of the hose must not be damaged. Lubricant OELPAG46 can be used to avoid mechanical abrasion; dry mounting is not allowed. Mind the marking of the insert depth on the hose cover.",
      "Push the ferrule forward to the insert nut.",
      "Swage as recommended.",
      "After swaging: visually control the assembly, control the crimp diameter, and control the hose cover marking of the insert depth (a small gap due to material deformation is tolerated).",
    ],
    onePieceWarningText: "A too large gap between the ferrule and the hose mark points to a wrong assembly: the interlock area of the ferrule isn't properly seated on the spiral braid.",
    onePieceNoteText: "Using a pre-assembling machine facilitates a proper assembly. A frequent failure cause of interlock fittings is a too large internal skive length.",
    onePieceLinkText: "This combination is assembled according to the One Piece Fitting instruction.",
    btnViewOnePiece: "View assembly instruction",

    // Interlock Fitting - reference page, based on HANSA-FLEX assembly instruction for interlock fittings (swaging chart p.30)
    interlockPageTitle: "Interlock Fitting — assembly instruction",
    interlockPageIntro: "Follow these steps carefully: incorrect assembly remains the most frequent cause of failure for interlock fittings.",
    interlockSteps: [
      "After skiving (internal and external): slide the ferrule onto the hose.",
      "Check the correct position of the hose relative to the ferrule collar.",
      "Mark the correct end position of the ferrule on the hose cover.",
      "Before swaging, insert the fitting up to the ferrule collar.",
      "Check the position of the ferrule, i.e. its alignment with the marking on the hose cover.",
      "Swage as recommended.",
      "After swaging, check the ferrule position against the hose mark; a small gap due to material deformation is acceptable.",
    ],
    interlockWarningText: "A too large gap between the ferrule and the hose mark points to a wrong assembly: the interlock area of the ferrule isn't properly seated on the spiral braid.",
    interlockNoteText: "Using a pre-assembling machine facilitates a proper assembly. A frequent failure cause of interlock fittings is a too large internal skive length.",

    // Stage Crimping - reference page, based on the HANSA-FLEX swaging chart (steel p.45, VA p.21):
    // combinations with assemblyType "stage" must be crimped to their final size in two or three steps.
    stageCrimpPageTitle: "Stage Crimping",
    stageCrimpPageIntro: "For a number of hose/fitting combinations (for example SGB 100 and NY 1800, in both steel and stainless), the final size must not be crimped in a single step. Work through the crimp diameters in the order shown; reaching the last (smallest) diameter gives the optimal connection.",
    stageCrimpWhyHeading: "Why in stages?",
    stageCrimpWhyText: "For these combinations, the difference between the uncrimped and final diameter is too large to crimp in one motion without damaging the hose reinforcement. Crimping via one or two intermediate sizes builds up the deformation gradually.",
    stageCrimpExampleHeading: "Example: SGB 100, DN 20",
    stageCrimpExampleText: "Ø1 = first (largest) crimp diameter, Ø2 = intermediate size, Ø3 = final size. See this combination's card for the exact values.",
    stageCrimpApplicableHeading: "Applies to",
    stageCrimpApplicableText: "SGB 100 (steel and stainless) and NY 1800. Combinations with this crimp schedule are marked with the \"Stage crimp\" label on the card.",
    stageCrimpLinkText: "This combination is crimped in stages.",
    btnViewStageCrimp: "View stage crimping instructions",

    // Stainless Fittings - reference page, based on the HANSA-FLEX crimp chart VA (05.2022).
    // Material rules that apply once a combination has material "rvs".
    stainlessPageTitle: "Stainless Fittings — material rules",
    stainlessPageIntro: "For stainless (VA) hose/fitting combinations, two extra material rules apply on top of the normal press instructions. Combinations with these rules are marked with the \"Stainless\" label on the card.",
    stainlessHoldTimeHeading: "Longer hold time in the press",
    stainlessHoldTimeText: "The hold time of stainless steel fittings in the press after reaching the crimp diameter has to be longer than for steel (approximately 5 seconds), to guarantee sufficient material deformation.",
    stainlessDisclaimerHeading: "Swaging diameters are standard values",
    stainlessDisclaimerText: "Swaging diameters are only standard values. The bore collapse must always be controlled, for example with the test mandrel; see also the internal works standard, chapter 8.",
    stainlessInterlockDiffHeading: "Difference in interlock assembly",
    stainlessInterlockDiffText: "For stainless interlock fittings, after sliding the ferrule up to the first resistance (the \"bump\") of the interlock, this resistance must deliberately be overcome. Then pull the ferrule slightly forward again to close the gap to the fitting, before crimping.",
    stainlessAutoNoteText: "Note: this is a stainless steel fitting. The hold time in the press has to be longer (approx. 5 sec.).",
    stainlessLinkText: "The stainless material rules apply to this combination.",
    btnViewStainlessInfo: "View stainless material rules",

    // Edit (add / edit / delete combinations)
    btnEdit: "Edit",
    btnAddCombination: "New combination",
    editTitleNew: "Add new combination",
    editTitleEdit: "Edit combination",
    editFieldSeries: "Hose type / series",
    editFieldStandard: "Standard",
    editFieldUitvoering: "Version",
    editFieldDn: "Nominal diameter (DN)",
    editFieldMfrCodes: "Manufacturer code(s)",
    editFieldMfrAdd: "Add code and press Enter",
    editFieldFerrule: "Ferrule / die",
    editFieldCrimp: "Crimp diameter (mm)",
    editFieldTol: "Tolerance (optional)",
    editFieldMandrel: "Mandrel (optional)",
    editFieldD1: "D1 (mm)",
    editFieldDa: "DA (mm, optional)",
    editFieldLf: "LF (mm)",
    editFieldSkive: "Skive length external (mm, optional)",
    editFieldSkiveInt: "Skive length internal (mm, optional)",
    editFieldAssemblyType: "Assembly type",
    assemblyTypeStandard: "Standard",
    assemblyTypeInterlock: "One Piece Fitting (interlock)",
    assemblyTypeStage: "Stage crimping",
    editFieldMaterial: "Material",
    materialLabelStaal: "Steel",
    materialLabelRvs: "Stainless (VA)",
    editFieldCrimp2: "Crimp diameter Ø2 (mm, for stage crimping)",
    editFieldCrimp3: "Crimp diameter Ø3 (mm, optional)",
    editFieldNote: "Note (optional)",
    btnSaveEdit: "Save",
    btnCancelEdit: "Cancel",
    btnDeleteItem: "Delete combination",
    btnDeleteConfirm: "Click again to delete permanently",
    editErrorRequired: "Fill in all required fields (series, standard, DN, manufacturer code, ferrule, crimp diameter, D1, LF).",
    editSavedRev: (revDate, seq) => `Saved as Rev. ${revDate}-R${String(seq).padStart(2, "0")}.`,

    updateTitle: "Update data",
    storageWarning: "Shared storage is not available in this environment. An upload will only work temporarily, for this screen.",
    introShared: "Upload an .xlsx file with the updated crimping data (use the template below as a starting point). After validation and confirmation, the data is stored in shared storage — everyone who opens this tool will then immediately see the new dataset. Revision numbers are updated automatically.",
    introLocal: "Upload an .xlsx file with the updated crimping data (use the template below as a starting point). After validation and confirmation, the data will be used temporarily on this screen only. Revision numbers are updated automatically.",
    btnChooseFile: "Choose file (.xlsx)",
    btnDownloadTemplate: "Download template",
    btnDownloadData: "Download current database",
    errFileType: "Only .xlsx files are supported.",
    btnAnalyse: "Analyse & validate",
    errorsFoundTitle: (n) => `${n} problem(s) found`,
    moreErrors: (n) => `… and ${n} more.`,
    validatedCountSuffix: "combination(s) validated, ready to save.",
    tableHeaderSeries: "Series",
    tableHeaderDn: "DN",
    tableHeaderHose: "Ferrule",
    moreRows: (n) => `… and ${n} more row(s).`,
    btnCancel: "Cancel",
    btnConfirmSave: "Confirm and save",
    savingLabel: "Saving…",
    statusUpdatedOn: (date, count) => `Current shared data updated on ${date} (${count} rows).`,
    statusNoUpload: "No shared upload is active yet — the tool is using the built-in sample data.",
    btnResetDefault: "Reset to sample data",
    btnResetConfirm: "Click again to confirm",
    fileReadError: "Could not read the file.",
    saveFailedDefault: "Saving failed.",

    errEmptyInput: "No data found — the xlsx file is empty.",
    errNoDataRows: "No data rows found.",
    errMissingColumns: (list) => `Required column(s) missing: ${list}.`,
    errSeriesMissing: (row) => `Row ${row}: 'series' is missing.`,
    errStandardMissing: (row) => `Row ${row}: 'standard' is missing.`,
    errUitvoeringInvalid: (row, found) => `Row ${row}: 'uitvoering' must be 'standaard' or 'geskived' (found: "${found}").`,
    errDnInvalid: (row) => `Row ${row}: 'dn' is not a valid number.`,
    errMfrCodesMissing: (row) => `Row ${row}: 'mfrCodes' is missing (separate multiple codes with a |, e.g. "25|27|35").`,
    errFerruleMissing: (row) => `Row ${row}: 'ferrule' is missing.`,
    errCrimpInvalid: (row) => `Row ${row}: 'crimp' is not a valid number.`,
    errD1Invalid: (row) => `Row ${row}: 'd1' is not a valid number.`,
    errLfInvalid: (row) => `Row ${row}: 'lf' is not a valid number.`,
    errRevMissing: (row) => `Row ${row}: 'rev' is missing.`,
    errRevDateMissing: (row) => `Row ${row}: 'revDate' is missing.`,

    printTitle: "HANSA-FLEX Crimp Data — Reference Card",
    printMetaLine: (label, date) => `Concept, sample data — ${label} — generated on ${date}`,
    printLabelOne: "1 combination",
    printLabelFavorites: "favourites",
    printLabelResults: "search results",
    printThSeries: "Series",
    printThMfrCode: "Mfr. code",
    printThHose: "Ferrule",
    printThMandrel: "Mandrel",
    printThD1LF: "D1 / LF",
    printThReference: "Reference",
    printInstructionsHeading: "Press instructions",
  },

  de: {
    appTitle: "Crimp Data",
    appSubtitle: "Digitale Presstabelle — Konzept",
    disclaimerStrong: "Konzeptdemo, Beispieldaten.",
    disclaimerRest: "Keine gültigen Produktionsspezifikationen. Bitte immer die offiziellen HANSA-FLEX-Presstabellen konsultieren.",
    btnUpdateData: "Daten aktualisieren",
    btnPrintExport: "Drucken / Exportieren",
    statusShared: (date, count) => `Gemeinsame Daten aktiv — aktualisiert am ${date} (${count} Zeilen)`,
    statusDefault: "Eingebaute Beispieldaten — noch kein Upload aktiv",
    tabSearch: "Suchen",
    tabFavorites: "Favoriten",
    tabMandrel: "Test-Dorn",
    tabOnePiece: "One Piece Fitting",
    tabInterlock: "Interlock Fitting",
    tabStageCrimp: "Stufenweise Verpressen",
    tabStainless: "Edelstahl-Armaturen",
    languageLabel: "Sprache",

    filterSeriesLabel: "Schlauchtyp / Baureihe",
    filterSeriesPlaceholder: "Z. B. HD 100, KP 200, TE 100 …",
    filterMfrLabel: "Schlauchcode",
    filterAll: "Alle",
    filterDnLabel: "Nennweite (DN)",
    filterVariantLabel: "Ausführung",
    variantStandardChip: "Standard (ohne Skive)",
    variantSkivedChip: "Geskivt",
    variantStandardShort: "ohne Skive",
    variantSkivedShort: "geskivt",
    resultsCount: (n) => `${n} Ergebnisse gefunden`,
    btnClearFilters: "Filter zurücksetzen",

    emptyResultsTitle: "Keine Kombinationen gefunden",
    emptyResultsMessage: "Passen Sie die Filter an oder setzen Sie sie zurück, um alle Beispieldaten anzuzeigen.",
    emptyFavoritesTitle: "Noch keine Favoriten",
    emptyFavoritesMessage: "Tippen Sie auf den Stern bei einem Ergebnis, um es hier schnell wiederzufinden.",
    favoritesCount: (n) => `${n} favorisierte Kombination(en)`,

    labelHose: "Hülse:",
    labelCrimpO: "Crimp-Ø",
    fixedSizeViaDie: "feste Maßeinstellung über Presswerkzeug",
    toleranceLabel: (tol) => `Toleranz ${tol}`,
    labelMandrel: "Dorn",
    mfrChipPrefix: "Herst.",
    fittingTypeV: "Typ V (Voss)",
    fittingTypeL: "Typ L (Lagra)",
    materialChipRvs: "Edelstahl",
    stageChipLabel: "Stufenweise",
    revShort: (revDate, seq) => `Rev. ${revDate}-R${String(seq || 1).padStart(2, "0")}`,
    btnDetails: "Details & Pressanleitung",
    favAriaSave: "Als Favorit speichern",
    favAriaRemove: "Aus Favoriten entfernen",
    noteIndicatorAria: "Enthält eine Anmerkung, siehe Details",

    fieldNominalDiameter: "Nennweite",
    fieldFerruleMarking: "Hülsenkennzeichnung",
    fieldMandrel: "Dorn",
    fieldD1LF: "D1 / LF",
    fieldDA: "DA (Hülsenaußendurchmesser)",
    fieldMfrCodes: "Herstellercode(s)",
    notApplicable: "entf.",
    headingInstructions: "Pressanleitung",
    referenceLine: (revDate, seq) => `Referenz: Revision Rev. ${revDate}-R${String(seq || 1).padStart(2, "0")}. Quelle: interne Presstabelle (Beispieldaten).`,
    btnSaveFavorite: "Als Favorit speichern",
    btnRemoveFavorite: "Favorit entfernen",
    btnPrintCard: "Karte drucken",
    ariaClose: "Schließen",

    stepMandrel: (mandrel, ferrule, dn) => `Dorn Typ ${mandrel} auswählen, passend zur Hülse ${ferrule} (DN${dn}).`,
    stepDieTable: (ferrule, dn) => `Presswerkzeug gemäß Presswerkzeugtabelle für Hülse ${ferrule} (DN${dn}) wählen.`,
    stepSkive: (len, intLen) => intLen ? `Schlauch außen auf ${len} mm und innen auf ${intLen} mm skivieren.` : `Schlauch (außen, und falls zutreffend innen) skivieren auf eine Länge von ${len} mm.`,
    stepSlideFerrule: "Hülse vollständig bis zum Anschlag auf den Schlauch schieben und die Position gegenüber der Endmarkierung kennzeichnen bzw. prüfen.",
    stepSetDiameter: (crimp, tol) =>
      `Crimp-/Pressdurchmesser auf ${crimp} mm einstellen${tol ? ` (Toleranz ${tol})` : " (feste Maßeinstellung über Presswerkzeug — genaue Toleranz siehe offizielle Tabelle)"}.`,
    stepSetDiameterStage: (diameters) =>
      `Stufenweise verpressen: zuerst Ø1 = ${diameters[0]} mm erreichen${diameters.length > 2 ? `, dann auf Ø2 = ${diameters[1]} mm reduzieren und schließlich auf Ø3 = ${diameters[2]} mm` : `, dann auf Ø2 = ${diameters[1]} mm reduzieren`}. Der letzte (kleinste) Durchmesser ergibt die optimale Verbindung, siehe auch den Reiter "Stufenweise Verpressen".`,
    stepPress: "Die Armatur in einem Arbeitsgang gemäß eingestelltem Durchmesser verpressen.",
    stepVerify: "Presswert mit mindestens 4 über den Umfang verteilten Messungen prüfen (z. B. bei 0°, 90°, 180° und 45°/135°).",
    stepVisualCheck: "Visuell auf Knicke oder Schiefstand prüfen; bei Zweifel die offizielle Tabelle für die Durchführung einer Druckprüfung konsultieren.",

    // Test-Dorn - Referenzseite, basierend auf HANSA-FLEX Werksnorm PN-DORN…
    mandrelPageTitle: "Test-Dorn",
    mandrelPageWarning: "Den Test-Dorn nur verwenden, wenn dies in der Presstabelle angegeben ist.",
    mandrelPageIntro: "Die HF-Test-Dorne dienen zur Kontrolle des Pressdurchmessers bei möglichen Toleranzabweichungen des Schlauchs. In der HF-Presstabelle ist angegeben, ob eine Bohrungseinschnürung (Bore Collapse) kontrolliert werden kann und mit welchem Durchmesser.",
    mandrelFigures: [
      { title: "Abbildung 1: korrekte Montage Ø A", text: "Eine Bohrungseinschnürung wurde erreicht. Der Durchmesser A des Test-Dorns kann bis zur Verpressung eingeführt werden, jedoch nicht weiter als die sogenannte \"Stufe 1\". Die Verpressung ist korrekt." },
      { title: "Abbildung 2: korrekte Montage Ø B", text: "Eine Bohrungseinschnürung wurde erreicht. Der Durchmesser B des Test-Dorns kann bis zur Verpressung eingeführt werden, jedoch nicht weiter als die sogenannte \"Stufe 2\". Die Verpressung ist korrekt." },
      { title: "Abbildung 3: unzureichende Einschnürung", text: "Der Test-Dorn geht vollständig in die Armatur. Die Bohrungseinschnürung ist unzureichend. Die Montage ist nicht korrekt. Vorgehensweise: den Pressdurchmesser in Schritten von 0,1 mm verringern, bis eine ausreichende Einschnürung erreicht ist (siehe Abbildung 1 oder 2)." },
      { title: "Abbildung 4: übermäßige Einschnürung", text: "Die Bohrung ist übermäßig eingeschnürt. Die extreme Verengung der Innenbohrung verhindert das Einführen des Dorns. Die Montage ist nicht korrekt; die Schlauchleitung darf nicht verwendet werden. Vorgehensweise: Schlauch- und Armaturmaße prüfen und den Pressdurchmesser anpassen, bis eine korrekte Montage erreicht ist (siehe Abbildung 1 oder 2)." },
    ],
    mandrelFrequencyHeading: "Prüfhäufigkeit",
    mandrelFrequencyText: "Bei Einzelfertigung wird jede Schlauchleitung geprüft. Bei Serienfertigung werden die ersten 6 Schlauchleitungen geprüft; danach erfolgen Stichproben, mindestens jede 20. Leitung, sowie immer bei einem Chargenwechsel des Schlauchs.",
    mandrelControlHeading: "Kontrolle des Test-Dorns",
    mandrelControlText: "Den Test-Dorn vor jedem Gebrauch visuell prüfen. Verdrehte oder gebrochene Test-Dorne müssen ersetzt werden. Eine leichte Abnutzung des Durchmessers A oder B ist unkritisch. Bei starker Beschädigung muss der Dorn sofort ersetzt werden.",

    // One Piece Fitting - Referenzseite, basierend auf HANSA-FLEX Montageanleitung PA 700 PLUS / HD 700 PLUS / HD 700 LL
    onePiecePageTitle: "One Piece Fitting — Montageanleitung",
    onePiecePageIntro: "Montageanleitung für Interlock-Armaturen (One Piece Fitting) wie PA 700 PLUS, HD 700 PLUS und HD 700 LL. Gilt für alle Kombinationen mit diesem Armaturtyp; siehe auch die Presskarte des jeweiligen Schlauchs.",
    onePieceSteps: [
      "Nach dem (innen- und außenseitigen) Skivieren die Einstecktiefe des Schlauchs aus der Presstabelle entnehmen und auf dem Schlauchmantel markieren. Diese Markierung hilft sicherzustellen, dass der Schlauch vollständig bis zur Endposition auf Armatur/Einsatz aufgeschoben ist.",
      "Die Hülse aufschieben.",
      "Achtung: bei Armaturen/Einsätzen in NW 32, 40 und 50 die 2 O-Ringe auf die dafür vorgesehenen Muttern am schlauchseitigen Ende setzen. Die korrekte Montage vor dem Zusammenbau visuell prüfen. Armaturen/Einsätze ohne diese 2 O-Ringe dürfen nicht montiert werden. Die Armatur/den Einsatz einsetzen (manuell oder mit einer Einsteckmaschine); das Innenrohr des Schlauchs darf nicht beschädigt werden. Das Schmiermittel OELPAG46 kann verwendet werden, um mechanischen Abrieb zu vermeiden; Trockenmontage ist nicht zulässig. Auf die Markierung der Einstecktiefe auf dem Schlauchmantel achten.",
      "Die Hülse bis zur Einsatzmutter vorschieben.",
      "Wie empfohlen verpressen.",
      "Nach dem Verpressen: die Montage visuell kontrollieren, den Crimp-Durchmesser kontrollieren und die Markierung der Einstecktiefe auf dem Schlauchmantel kontrollieren (ein kleiner Spalt durch Materialverformung ist zulässig).",
    ],
    onePieceWarningText: "Ein zu großer Spalt zwischen Hülse und Schlauchmarkierung deutet auf eine fehlerhafte Montage hin: der Interlock-Bereich der Hülse liegt dann nicht richtig auf dem Spiralgeflecht.",
    onePieceNoteText: "Die Verwendung einer Vormontagemaschine erleichtert eine korrekte Montage. Eine häufige Fehlerursache bei Interlock-Armaturen ist eine zu große innere Skivelänge.",
    onePieceLinkText: "Diese Kombination wird gemäß der One-Piece-Fitting-Anleitung montiert.",
    btnViewOnePiece: "Montageanleitung ansehen",

    // Interlock Fitting - Referenzseite, basierend auf HANSA-FLEX Montageanleitung für Interlock-Armaturen (Presstabelle S.30)
    interlockPageTitle: "Interlock Fitting — Montageanleitung",
    interlockPageIntro: "Beachten Sie diese Schritte sorgfältig: eine fehlerhafte Montage bleibt die häufigste Fehlerursache bei Interlock-Armaturen.",
    interlockSteps: [
      "Nach dem Skivieren (innen und außen): Schieben Sie die Hülse auf den Schlauch.",
      "Kontrollieren Sie die korrekte Position von Schlauch und Hülsenkragen (Collar).",
      "Markieren Sie die korrekte Endposition der Hülse auf dem Schlauchmantel.",
      "Stecken Sie vor dem Pressen die Armatur/den Einsatz bis zum Hülsenkragen ein.",
      "Kontrollieren Sie die Position der Hülse, das heißt die Übereinstimmung mit der Markierung auf dem Schlauchmantel.",
      "Pressen Sie wie empfohlen.",
      "Kontrollieren Sie nach dem Pressen die Hülsenposition gegenüber der Schlauchmarkierung; ein geringer Spalt durch Materialverformung ist zulässig.",
    ],
    interlockWarningText: "Ein zu großer Spalt zwischen Hülse und Schlauchmarkierung deutet auf eine fehlerhafte Montage hin: der Interlock-Bereich der Hülse liegt dann nicht richtig auf dem Spiralgeflecht.",
    interlockNoteText: "Die Verwendung einer Vormontagemaschine erleichtert eine korrekte Montage. Eine häufige Fehlerursache bei Interlock-Armaturen ist eine zu große innere Skivelänge.",

    // Stufenweise Verpressen - Referenzseite, basierend auf der HANSA-FLEX Presstabelle (Stahl S.45, VA S.21):
    // Kombinationen mit assemblyType "stage" müssen in zwei oder drei Schritten auf die Endmaß verpresst werden.
    stageCrimpPageTitle: "Stufenweise Verpressen",
    stageCrimpPageIntro: "Bei einigen Schlauch-Armatur-Kombinationen (z. B. SGB 100 und NY 1800, sowohl in Stahl als auch in Edelstahl) darf das Endmaß nicht in einem Schritt verpresst werden. Die Crimp-Durchmesser in der angegebenen Reihenfolge durchlaufen; beim Erreichen des letzten (kleinsten) Durchmessers ist die Verbindung optimal.",
    stageCrimpWhyHeading: "Warum stufenweise?",
    stageCrimpWhyText: "Bei diesen Kombinationen ist der Sprung zwischen dem unverpressten und dem endgültigen Durchmesser zu groß, um ihn in einer Bewegung zu verpressen, ohne die Schlaucharmierung zu beschädigen. Durch das Verpressen über eine oder zwei Zwischenstufen wird die Verformung schrittweise aufgebaut.",
    stageCrimpExampleHeading: "Beispiel: SGB 100, DN 20",
    stageCrimpExampleText: "Ø1 = erster (größter) Pressdurchmesser, Ø2 = Zwischenmaß, Ø3 = Endmaß. Die genauen Werte siehe Karte dieser Kombination.",
    stageCrimpApplicableHeading: "Gilt für",
    stageCrimpApplicableText: "SGB 100 (Stahl und Edelstahl) und NY 1800. Kombinationen mit diesem Pressschema sind an der Kennzeichnung \"Stufenweise\" auf der Karte erkennbar.",
    stageCrimpLinkText: "Diese Kombination wird stufenweise verpresst.",
    btnViewStageCrimp: "Stufenweise Pressanleitung ansehen",

    // Edelstahl-Armaturen - Referenzseite, basierend auf der HANSA-FLEX crimp chart VA (05.2022).
    // Materialregeln, die gelten, sobald eine Kombination material "rvs" hat.
    stainlessPageTitle: "Edelstahl-Armaturen — Materialregeln",
    stainlessPageIntro: "Für Edelstahl (VA) Schlauch-Armatur-Kombinationen gelten, zusätzlich zur normalen Pressanleitung, zwei weitere Materialregeln. Kombinationen mit diesen Regeln sind an der Kennzeichnung \"Edelstahl\" auf der Karte erkennbar.",
    stainlessHoldTimeHeading: "Längere Haltezeit in der Presse",
    stainlessHoldTimeText: "Die Haltezeit von Edelstahl-Armaturen in der Presse muss nach Erreichen des Crimp-Durchmessers länger sein als bei Stahl (ca. 5 Sekunden), um eine ausreichende Materialverformung zu gewährleisten.",
    stainlessDisclaimerHeading: "Pressmaße sind Standardwerte",
    stainlessDisclaimerText: "Pressdurchmesser sind ausschließlich Standardwerte. Der Bohrungskollaps muss immer kontrolliert werden, z. B. mit dem Test-Dorn; siehe auch die interne Werknorm Kapitel 8.",
    stainlessInterlockDiffHeading: "Abweichung bei der Interlock-Montage",
    stainlessInterlockDiffText: "Bei Edelstahl-Interlock-Armaturen muss nach dem Aufschieben der Hülse bis zum ersten Widerstand (dem \"Bump\") des Interlocks dieser Widerstand bewusst überwunden werden. Danach die Hülse wieder etwas nach vorne ziehen, um den Spalt zur Armatur zu schließen, bevor verpresst wird.",
    stainlessAutoNoteText: "Achtung: dies ist eine Edelstahl-Armatur. Die Haltezeit in der Presse muss länger sein (ca. 5 Sek.).",
    stainlessLinkText: "Für diese Kombination gelten die Edelstahl-Materialregeln.",
    btnViewStainlessInfo: "Edelstahl-Materialregeln ansehen",

    // Wijzigen (Kombinationen bearbeiten / hinzufügen / löschen)
    btnEdit: "Bearbeiten",
    btnAddCombination: "Neue Kombination",
    editTitleNew: "Neue Kombination hinzufügen",
    editTitleEdit: "Kombination bearbeiten",
    editFieldSeries: "Schlauchtyp / Baureihe",
    editFieldStandard: "Norm",
    editFieldUitvoering: "Ausführung",
    editFieldDn: "Nennweite (DN)",
    editFieldMfrCodes: "Herstellercode(s)",
    editFieldMfrAdd: "Code eingeben und Enter drücken",
    editFieldFerrule: "Hülse / Presswerkzeug",
    editFieldCrimp: "Crimp-Durchmesser (mm)",
    editFieldTol: "Toleranz (optional)",
    editFieldMandrel: "Dorn (optional)",
    editFieldD1: "D1 (mm)",
    editFieldDa: "DA (mm, optional)",
    editFieldLf: "LF (mm)",
    editFieldSkive: "Skivelänge außen (mm, optional)",
    editFieldSkiveInt: "Skivelänge innen (mm, optional)",
    editFieldAssemblyType: "Montageart",
    assemblyTypeStandard: "Standard",
    assemblyTypeInterlock: "One Piece Fitting (Interlock)",
    assemblyTypeStage: "Stufenweise Verpressen",
    editFieldMaterial: "Material",
    materialLabelStaal: "Stahl",
    materialLabelRvs: "Edelstahl (VA)",
    editFieldCrimp2: "Crimp-Durchmesser Ø2 (mm, bei stufenweisem Verpressen)",
    editFieldCrimp3: "Crimp-Durchmesser Ø3 (mm, optional)",
    editFieldNote: "Anmerkung (optional)",
    btnSaveEdit: "Speichern",
    btnCancelEdit: "Abbrechen",
    btnDeleteItem: "Kombination löschen",
    btnDeleteConfirm: "Erneut klicken, um endgültig zu löschen",
    editErrorRequired: "Bitte alle Pflichtfelder ausfüllen (Baureihe, Norm, DN, Herstellercode, Hülse, Crimp-Durchmesser, D1, LF).",
    editSavedRev: (revDate, seq) => `Gespeichert als Rev. ${revDate}-R${String(seq).padStart(2, "0")}.`,

    updateTitle: "Daten aktualisieren",
    storageWarning: "Gemeinsamer Speicher ist in dieser Umgebung nicht verfügbar. Ein Upload funktioniert jetzt nur vorübergehend, nur für diesen Bildschirm.",
    introShared: "Laden Sie eine .xlsx-Datei mit den aktualisierten Crimpdaten hoch (nutzen Sie die Vorlage unten als Ausgangspunkt). Nach Prüfung und Bestätigung werden die Daten gemeinsam gespeichert — alle, die dieses Tool öffnen, sehen danach sofort den neuen Datensatz. Revisionsnummern werden automatisch aktualisiert.",
    introLocal: "Laden Sie eine .xlsx-Datei mit den aktualisierten Crimpdaten hoch (nutzen Sie die Vorlage unten als Ausgangspunkt). Nach Prüfung und Bestätigung werden die Daten nur vorübergehend auf diesem Bildschirm verwendet. Revisionsnummern werden automatisch aktualisiert.",
    btnChooseFile: "Datei auswählen (.xlsx)",
    btnDownloadTemplate: "Vorlage herunterladen",
    btnDownloadData: "Aktuelle Datenbank herunterladen",
    errFileType: "Es werden nur .xlsx-Dateien unterstützt.",
    btnAnalyse: "Analysieren & prüfen",
    errorsFoundTitle: (n) => `${n} Problem(e) gefunden`,
    moreErrors: (n) => `… und ${n} weitere.`,
    validatedCountSuffix: "Kombination(en) validiert, bereit zum Speichern.",
    tableHeaderSeries: "Baureihe",
    tableHeaderDn: "DN",
    tableHeaderHose: "Hülse",
    moreRows: (n) => `… und ${n} weitere Zeile(n).`,
    btnCancel: "Abbrechen",
    btnConfirmSave: "Bestätigen und speichern",
    savingLabel: "Wird gespeichert…",
    statusUpdatedOn: (date, count) => `Aktuelle gemeinsame Daten aktualisiert am ${date} (${count} Zeilen).`,
    statusNoUpload: "Es ist noch kein gemeinsamer Upload aktiv — das Tool verwendet die eingebauten Beispieldaten.",
    btnResetDefault: "Auf Beispieldaten zurücksetzen",
    btnResetConfirm: "Erneut klicken zum Bestätigen",
    fileReadError: "Datei konnte nicht gelesen werden.",
    saveFailedDefault: "Speichern fehlgeschlagen.",

    errEmptyInput: "Keine Daten gefunden — die xlsx-Datei ist leer.",
    errNoDataRows: "Keine Datenzeilen gefunden.",
    errMissingColumns: (list) => `Pflichtspalte(n) fehlen: ${list}.`,
    errSeriesMissing: (row) => `Zeile ${row}: 'series' fehlt.`,
    errStandardMissing: (row) => `Zeile ${row}: 'standard' fehlt.`,
    errUitvoeringInvalid: (row, found) => `Zeile ${row}: 'uitvoering' muss 'standaard' oder 'geskived' sein (gefunden: "${found}").`,
    errDnInvalid: (row) => `Zeile ${row}: 'dn' ist keine gültige Zahl.`,
    errMfrCodesMissing: (row) => `Zeile ${row}: 'mfrCodes' fehlt (mehrere Codes mit | trennen, z. B. "25|27|35").`,
    errFerruleMissing: (row) => `Zeile ${row}: 'ferrule' fehlt.`,
    errCrimpInvalid: (row) => `Zeile ${row}: 'crimp' ist keine gültige Zahl.`,
    errD1Invalid: (row) => `Zeile ${row}: 'd1' ist keine gültige Zahl.`,
    errLfInvalid: (row) => `Zeile ${row}: 'lf' ist keine gültige Zahl.`,
    errRevMissing: (row) => `Zeile ${row}: 'rev' fehlt.`,
    errRevDateMissing: (row) => `Zeile ${row}: 'revDate' fehlt.`,

    printTitle: "HANSA-FLEX Crimp Data — Referenzkarte",
    printMetaLine: (label, date) => `Konzept, Beispieldaten — ${label} — erstellt am ${date}`,
    printLabelOne: "1 Kombination",
    printLabelFavorites: "Favoriten",
    printLabelResults: "Suchergebnisse",
    printThSeries: "Baureihe",
    printThMfrCode: "Herst.-Code",
    printThHose: "Hülse",
    printThMandrel: "Dorn",
    printThD1LF: "D1 / LF",
    printThReference: "Referenz",
    printInstructionsHeading: "Pressanleitung",
  },
};

function variantChipLabel(value, t) {
  return value === "geskived" ? t.variantSkivedChip : t.variantStandardChip;
}
function variantShortLabel(uitvoering, t) {
  return uitvoering === "geskived" ? t.variantSkivedShort : t.variantStandardShort;
}
function itemNote(item, lang) {
  const translated = NOTE_TRANSLATIONS[item.id];
  if (translated && translated[lang]) return translated[lang];
  return item.note;
}

// Taalonafhankelijke check of een combinatie een opmerking heeft (voor het
// waarschuwingsicoontje op de resultaatkaart, dat niet van de actieve taal afhangt).
function hasNote(item) {
  return !!(item.note || NOTE_TRANSLATIONS[item.id]);
}

// Voor interlock-combinaties met een Typ "V" (Voss) / Typ "L" (Lagra) huls in de
// hulsaanduiding: geeft het bijpassende label terug, zodat de gebruiker bij het
// kiezen tussen resultaten direct ziet welk fitting-merk/type hoort bij welke persmaat.
function fittingTypeLabel(item, t) {
  if (!item.ferrule) return null;
  if (item.ferrule.includes("Typ „V“")) return t.fittingTypeV;
  if (item.ferrule.includes("Typ „L“")) return t.fittingTypeL;
  return null;
}

/* ------------------------------------------------------------------ */
/*  Upload- en validatiehulpfuncties (voor "Data bijwerken")           */
/* ------------------------------------------------------------------ */
const REQUIRED_FIELDS = ["series", "standard", "uitvoering", "dn", "mfrCodes", "ferrule", "crimp", "d1", "lf"];

function slugify(s) {
  return String(s)
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/(^-|-$)/g, "");
}

function normalizeRow(raw, index, t) {
  const errors = [];
  const rowNr = index + 1;

  const series = (raw.series || "").toString().trim();
  if (!series) errors.push(t.errSeriesMissing(rowNr));

  const standard = (raw.standard || "").toString().trim();
  if (!standard) errors.push(t.errStandardMissing(rowNr));

  let uitvoering = (raw.uitvoering || "").toString().trim().toLowerCase();
  if (uitvoering !== "standaard" && uitvoering !== "geskived") {
    errors.push(t.errUitvoeringInvalid(rowNr, raw.uitvoering || ""));
    uitvoering = "standaard";
  }

  const dn = Number(raw.dn);
  if (raw.dn === undefined || raw.dn === "" || Number.isNaN(dn) || dn <= 0) {
    errors.push(t.errDnInvalid(rowNr));
  }

  let mfrCodes = raw.mfrCodes;
  if (Array.isArray(mfrCodes)) {
    mfrCodes = mfrCodes.map((c) => String(c).trim()).filter(Boolean);
  } else {
    mfrCodes = (mfrCodes || "").toString().split("|").map((c) => c.trim()).filter(Boolean);
  }
  if (mfrCodes.length === 0) errors.push(t.errMfrCodesMissing(rowNr));

  const ferrule = (raw.ferrule || "").toString().trim();
  if (!ferrule) errors.push(t.errFerruleMissing(rowNr));

  const crimp = Number(raw.crimp);
  if (raw.crimp === undefined || raw.crimp === "" || Number.isNaN(crimp) || crimp <= 0) {
    errors.push(t.errCrimpInvalid(rowNr));
  }

  const d1 = Number(raw.d1);
  if (raw.d1 === undefined || raw.d1 === "" || Number.isNaN(d1)) errors.push(t.errD1Invalid(rowNr));

  const lf = Number(raw.lf);
  if (raw.lf === undefined || raw.lf === "" || Number.isNaN(lf)) errors.push(t.errLfInvalid(rowNr));

  const tolRaw = (raw.tol ?? "").toString().trim();
  const tol = tolRaw ? tolRaw : null;

  const mandrelRaw = (raw.mandrel ?? "").toString().trim();
  const mandrel = mandrelRaw ? mandrelRaw : null;

  const note = raw.note ? String(raw.note).trim() : undefined;

  const skiveLengthRaw = raw.skiveLength;
  const skiveLength =
    skiveLengthRaw !== undefined && skiveLengthRaw !== "" && !Number.isNaN(Number(skiveLengthRaw))
      ? Number(skiveLengthRaw)
      : undefined;

  const skiveLengthIntRaw = raw.skiveLengthInt;
  const skiveLengthInt =
    skiveLengthIntRaw !== undefined && skiveLengthIntRaw !== "" && !Number.isNaN(Number(skiveLengthIntRaw))
      ? Number(skiveLengthIntRaw)
      : undefined;

  const daRaw = raw.da;
  const da = daRaw !== undefined && daRaw !== "" && !Number.isNaN(Number(daRaw)) ? Number(daRaw) : undefined;

  const assemblyTypeRaw = (raw.assemblyType || "").toString().trim().toLowerCase();
  const assemblyType = ["interlock", "custom", "stage"].includes(assemblyTypeRaw) ? assemblyTypeRaw : "standard";

  const materialRaw = (raw.material || "").toString().trim().toLowerCase();
  const material = materialRaw === "rvs" ? "rvs" : undefined;

  const crimp2Raw = raw.crimp2;
  const crimp2 = crimp2Raw !== undefined && crimp2Raw !== "" && !Number.isNaN(Number(crimp2Raw)) ? Number(crimp2Raw) : undefined;

  const crimp3Raw = raw.crimp3;
  const crimp3 = crimp3Raw !== undefined && crimp3Raw !== "" && !Number.isNaN(Number(crimp3Raw)) ? Number(crimp3Raw) : undefined;

  const idRaw = (raw.id || "").toString().trim();
  const id = idRaw || `${slugify(series) || "combi"}-dn${raw.dn || index}-${slugify(ferrule) || index}`;

  const row = {
    id,
    series,
    standard,
    uitvoering,
    uitvoeringLabel: uitvoering === "geskived" ? "geskived" : "geen skive",
    dn,
    mfrCodes,
    ferrule,
    crimp,
    tol,
    mandrel,
    d1,
    lf,
    assemblyType,
  };
  if (note) row.note = note;
  if (skiveLength) row.skiveLength = skiveLength;
  if (skiveLengthInt) row.skiveLengthInt = skiveLengthInt;
  if (da) row.da = da;
  if (material) row.material = material;
  if (crimp2 !== undefined) row.crimp2 = crimp2;
  if (crimp3 !== undefined) row.crimp3 = crimp3;

  return { row, errors };
}

// Valideert rijen (objecten, zoals die uit een xlsx-werkblad komen). Revisiegegevens
// (revDate/revSeq) horen hier niet bij — die worden automatisch beheerd bij het opslaan.
function validateRows(rawRows, t) {
  if (!rawRows || rawRows.length === 0) {
    return { rows: [], errors: [t.errEmptyInput] };
  }

  const missingHeader = REQUIRED_FIELDS.filter((f) => !(f in rawRows[0]));
  const errors = [];
  if (missingHeader.length > 0) {
    errors.push(t.errMissingColumns(missingHeader.join(", ")));
  }

  const rows = [];
  rawRows.forEach((raw, i) => {
    const { row, errors: rowErrors } = normalizeRow(raw, i, t);
    errors.push(...rowErrors);
    rows.push(row);
  });

  return { rows, errors };
}

// Leest een geüpload .xlsx-bestand (eerste werkblad) in als array van objecten,
// met de kolomkoppen uit de eerste rij als sleutels.
function readXlsxFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const data = new Uint8Array(ev.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        resolve(rows);
      } catch (e) {
        reject(e);
      }
    };
    reader.onerror = () => reject(new Error("read error"));
    reader.readAsArrayBuffer(file);
  });
}

function buildTemplateXlsx() {
  const header = ["id", "series", "standard", "uitvoering", "dn", "mfrCodes", "ferrule", "crimp", "crimp2", "crimp3", "tol", "mandrel", "d1", "da", "lf", "skiveLength", "skiveLengthInt", "assemblyType", "material", "note"];
  const example = [
    ["", "HD 100 - no skive - PHN 100", "EN 853 1SN", "standaard", 6, "25|27|35", "HF-xxx PHN106", 16.6, "", "", "", "B", 19.6, "", 30.2, "", "", "standard", "", ""],
    ["", "HD 100 - no skive - PHN 200", "EN 853 1SN", "standaard", 20, "25|27|35|50", "HF-xxx PHN220", 31.4, "", "", "± 0,1 mm", "", 37.0, "", 42.5, "", "", "standard", "", ""],
    ["", "HD 100 - no skive VA", "EN 853 1SN", "standaard", 6, "5|25|27|35", "HF-xxx PHD106VA", 16.7, "", "", "± 0,1 mm", "", 20.0, "", 34.5, "", "", "standard", "rvs", ""],
    ["", "SGB 100 - stage crimping", "exceeds SAE100R4", "standaard", 20, "25", "HF-xxx PHN220", 33.0, 32.5, 32.0, "", "", 37.0, "", 42.5, "", "", "stage", "", ""],
  ];
  const ws = XLSX.utils.aoa_to_sheet([header, ...example]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "crimpdata");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}

// Exporteert de volledig actieve dataset (alle velden, inclusief id), zodat deze na
// aanpassing weer via "Bestand kiezen" geupload kan worden. De laatste kolom toont
// altijd de huidige revisie (zoals ook op de kaart/detailweergave staat); deze kolom
// is alleen informatief en wordt bij upload genegeerd — de revisie wordt automatisch
// bepaald op basis van eventuele inhoudelijke wijzigingen.
function buildDataXlsx(rows) {
  const header = ["id", "series", "standard", "uitvoering", "dn", "mfrCodes", "ferrule", "crimp", "crimp2", "crimp3", "tol", "mandrel", "d1", "da", "lf", "skiveLength", "skiveLengthInt", "assemblyType", "material", "note", "revision"];
  const body = rows.map((r) => [
    r.id,
    r.series,
    r.standard,
    r.uitvoering,
    r.dn,
    Array.isArray(r.mfrCodes) ? r.mfrCodes.join("|") : "",
    r.ferrule,
    r.crimp,
    r.crimp2 != null ? r.crimp2 : "",
    r.crimp3 != null ? r.crimp3 : "",
    r.tol || "",
    r.mandrel || "",
    r.d1,
    r.da != null ? r.da : "",
    r.lf,
    r.skiveLength != null ? r.skiveLength : "",
    r.skiveLengthInt != null ? r.skiveLengthInt : "",
    r.assemblyType || "standard",
    r.material || "",
    r.note || "",
    `Rev. ${r.revDate}-R${String(r.revSeq || 1).padStart(2, "0")}`,
  ]);
  const ws = XLSX.utils.aoa_to_sheet([header, ...body]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "crimpdata");
  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}

function downloadBinaryFile(filename, arrayBuffer, mime) {
  const blob = new Blob([arrayBuffer], { type: mime || "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ------------------------------------------------------------------ */
/*  Hulpfuncties                                                      */
/* ------------------------------------------------------------------ */
// Revisiesysteem: elke combinatie draagt haar eigen "Rev. YYYY-MM-R0x". Bij een
// wijziging (handmatig via "Wijzigen" of via een xlsx-upload) wordt de teller
// opgehoogd binnen dezelfde maand, of teruggezet naar R01 in een nieuwe maand.
function currentRevMonth() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
}

function bumpRevision(existingItem) {
  const month = currentRevMonth();
  if (existingItem && existingItem.revDate === month) {
    return { revDate: month, revSeq: (existingItem.revSeq || 1) + 1 };
  }
  return { revDate: month, revSeq: 1 };
}

// Velden die de inhoud van een combinatie bepalen (id en revDate/revSeq horen hier
// niet bij: die zijn geen inhoudelijke wijziging). Gebruikt om bij een xlsx-heruploaded
// te bepalen of een rij daadwerkelijk gewijzigd is, zodat de revisie alleen wordt
// opgehoogd als er echt iets veranderd is.
const DIFF_FIELDS = ["series", "standard", "uitvoering", "dn", "mfrCodes", "ferrule", "crimp", "crimp2", "crimp3", "tol", "mandrel", "d1", "da", "lf", "skiveLength", "skiveLengthInt", "assemblyType", "material", "note"];

function normalizeForCompare(row) {
  const out = {};
  DIFF_FIELDS.forEach((f) => {
    let v = row ? row[f] : undefined;
    if (f === "mfrCodes") v = Array.isArray(v) ? [...v].map(String).sort().join("|") : "";
    else if (v === undefined || v === null || v === "") v = "";
    else v = String(v);
    out[f] = v;
  });
  return out;
}

function rowsContentEqual(a, b) {
  const na = normalizeForCompare(a);
  const nb = normalizeForCompare(b);
  return DIFF_FIELDS.every((f) => na[f] === nb[f]);
}

// Voor combinaties met assemblyType "stage": geeft de opeenvolgende persdiameters
// terug (Ø1, Ø2, en optioneel Ø3) als array van strings met 1 decimaal.
function stageCrimpDiameters(item) {
  const values = [item.crimp, item.crimp2, item.crimp3].filter((v) => v !== undefined && v !== null);
  return values.map((v) => v.toFixed(1));
}

function getSteps(item, t, lang) {
  if (item.assemblyType === "custom" && CUSTOM_STEPS[item.id]) {
    return CUSTOM_STEPS[item.id][lang] || CUSTOM_STEPS[item.id].nl;
  }
  const steps = [];
  if (item.mandrel) {
    steps.push(t.stepMandrel(item.mandrel, item.ferrule, item.dn));
  } else {
    steps.push(t.stepDieTable(item.ferrule, item.dn));
  }
  if (item.skiveLength) {
    steps.push(t.stepSkive(item.skiveLength, item.skiveLengthInt));
  }
  steps.push(t.stepSlideFerrule);
  if (item.assemblyType === "stage" && item.crimp2 != null) {
    steps.push(t.stepSetDiameterStage(stageCrimpDiameters(item)));
  } else {
    steps.push(t.stepSetDiameter(item.crimp.toFixed(1), item.tol));
  }
  steps.push(t.stepPress);
  steps.push(t.stepVerify);
  steps.push(t.stepVisualCheck);
  return steps;
}

/* ------------------------------------------------------------------ */
/*  Kleine subcomponenten                                             */
/* ------------------------------------------------------------------ */

const HF_LOGO_SRC =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAABvUAAADtCAYAAACPgfnQAAAACXBIWXMAAC4jAAAuIwF4pT92AAAgAElEQVR4nO3d3XEbx7Yo4JlbfqcQAXXe8HJB7QhIRyCeCERFIDoCyRGYikBUBJuKYEsRWMTLxdumIoAYwdxqu2lTNCUC4Pz0z/dVqbzPKdkc9IDAml691mq7rmsAAACAabRte9Q0zVXXdVduAQAA8D3/x8oAAADApM6apvlv27Yf27Y9disAAID7qNQDAACAibRt+6xpmt/v/PQvTdOch2Rf13Vf3RsAAKBRqQcAAACTOr3nh+83TfO6aZp127bnsT0nAABQOZV6AAAAMIG2bZ+EWXpN0+xt8NMvY5vOC9V7AABQJ5V6AAAAMI3jDRN6wUHTNO9CErBt27O2bZ+6ZwAAUBeVegAAADCBtm0/x2Tdrj7FuXsX7h8AAJRPUg8AAABG1rbts6Zpfu/pp35pmuY8Jvi05gQAgEJpvwkAAADjO+3xJ+43TfO6aZp127bnbdseuZ8AAFAelXoAAAAworZtn4TZeFvM09vFZajca5rmQvUeAACUQaUeAAAAjOt44IReE2f1vQvJw7Ztz9q2feoeAwBA3lTqAQAAwIjatv0ck25j+xRm73Vdd+5+AwBAfiT1AAAAYCRt2z5rmub3idf7S0juxQTf1cTXAgAAbEj7TQAAABjPSQJrvd80zeumaf7btu1527ZHCVwTAADwAJV6AAAAMJK2bb+OME9vF6F6703TNBdd1331fgAAgPSo1AMAAIARtG17kmhCr4nVe++aprmK1XtPE7gmAADgFkk9AAAAGMdxBuscko4vYmvOjzERCQAAJED7TQAAABhYrHz7b6brHFpznoc/XdddJXA9AABQJZV6AAAAMLwcqvS+J7TmfB2r947SvEQAACifSj0AAAAYWNu2VzE5lrPrpmmedl331fsFAADGp1IPAAAABtS27bMCEnrBhYQeAABMR1IPAAAAhnVayPqeJXANAABQLe03AQAAYEBt24bqtr3M1/iy67pnCVwHAABUS6UeAAAADKRt25MCEnqNKj0AAJiepB4AAAAM57iAtb0O8/QSuA4AAKiapB4AAAAMoG3bJ03TPC9gbS+6rvuawHUAAEDVJPUAAABgGCeFrKvWmwAAkABJPQAAABhGCUm9y67rPidwHQAAUD1JPQAAAOhZ27ZPm6Y5KGBdVekBAEAiJPUAAACgf6cFrOl1mKeXwHUAAED1Gkk9AAAAGMRxAct60XXd1wSuAwAAqtdI6gEAAEC/2rZ91jTNfgHLep7ANQAAAJGkHgAAAPTrpID1/NJ13ccErgMAAIgk9QAAAKBfJST1zqb84avZ4kn4M+U1AABAaiT1AAAAoCdt24ZZensFrOfFxD8/VAl+XM0WTye+DgAASIakHgAAAPTnuIC1/NB13dVUP3w1W4RZfgfxz+fVbPFsqmsBAICUSOoBAABAf0pI6k1WpRcTei9u/b/2YsWexB4AANVru66rfQ0AAADg0WLrzX9nvpLXXddNMsvunoTebddN0xzN18vPU1wbAACkQKUeAAAA9EOV3o4eSOg1tyr2Tqa4PgAASIFKPQAAAOhB27ZfY/IpZ//qum7UariYqHu3xb/ycr5eng94SQAAkCSVegAAAPBIsfVm7gm9Lxkk9IJ3KvYAAKiRpB4AAAA8XgmtN0etftsxoXcjJPZO+78qAABIl/abAAAA8EiFtN78n67rrsb4QavZ4lmYkdfDmr2fr5eq9gAAqIJKPQAAAHiEQlpvXmaY0AterGYL8/UAAKiCpB4AAAA8TgmtN8/G+CGr2eJJ0zQXPSdBJfYAAKiC9psAAADwCIW03px1Xfd1yB8QE3qhQu9goB/xKSRY5+vloK8DAACmolIPAAAAdlRI680PQyf0ovMBE3rBYUgaxuQhAAAUR1IPAAAAdldC682LoX9AbI/5fOifE5OGEnsAABRJ+00AAADYUQGtN6+bpnk6ZKXearY4aZrm3VD//e+4jK04r0b+uQAAMBiVegAAALCDQlpvXgyc0DueIKHXxIq9z6vZ4tkEPxsAAAYhqQcAAAC70XrzB2JC7Xyo//4G9mIrTok9AACKoP0mAAAA7KCE1ptd1w0yey7OtPvcNM3+EP/9LYUWo0fz9fJzAtcCAAA7U6kHAAAAW2rb9lkJrTcH/G9/TCSh19yq2DtJ4FoAAGBnknoAAACwvRISRIMk9VazxXmcaZeSkNh7J7EHAEDOtN8EAACALbVte5VQJdouBmm9GZNm70Z5Bbt7OV8vp5z1BwAAO1GpBwAAAFuIrTdzTug1Q1TprWaLZxkk9JpYsXeawHUAAMBWJPUAAABgO0cFrFevSb3VbPEkztHLxW+xTSgAAGRDUg8AAAC2k/tcttB6s+9KvY9xbl1OXkjsAQCQE0k9AAAA2FDbtk+bpjnIfL36rtI7z3hNJPYAAMiGpB4AAABsTuvNW1azRahafNHXf28iIbH3MbYQBQCAZLVd17k7AAAAsIG2bUNC7HnGaxVab/aSvFrNFs8ybbv5PZchaTtfL7+meXkAANROpR4AAABsLueEXtNXlV6sajsvKKHXxBaiKvYAAEiWpB4AAABsoG3b4wLWqa/WmznP0fuRm8Te03QvEQCAWknqAQAAwGZyT+qF1puPTuqtZovTAioWfyQk9j7H9qIAAJAMST0AAADYzFHm69RHQi8kun7r53KSthcr9iT2AABIhqQeAAAAPKBt25Dc2c98nR6V1Iuz5vpq35mDkNh7U9HrBQAgcZJ6AAAA8LDcq/SCj4/8988LSGxu47JpmpN8LhcAgNJJ6gEAAMDDcp+n96Hruq+7/sur2eKk8Dl6d4WE3tF8vdx5zQAAoG+SegAAAPADbduGtpOHma/Rzm0zV7PF06Zpzvq9nKRJ6AEAkCRJPQAAAPixElpvPmYW3kWcL1eDawk9AABSJakHAAAAP5Z7683LXVtvrmaLN03THPR/SUmS0AMAIGmSegAAAPBjuVfqne/yL61mi/C6X/d/OUm6Seh9ruT1AgCQobbrOvcNAAAA7tG27bOmaX7PfG3+p+u6q23+hdVsEeYIhgTX/nCXlZSf5+vlx0peKwAAmVKpBwAAAN+Xe5Xel20TetFZRQm9lxJ6AADkQFIPAAAAvi/3eXoX2/4Lq9kivOYXw1xOckJCb6f2pAAAMDbtNwEAAOA72rbN/aH5567rNq5Ci203Q2Xf3rCXlYT38/XypILXCQBAIVTqAQAAwD3ats29Su96m4RedC6hBwAAaZLUAwAAgPvlPk9vq7aSse3m8+EuJxmXTdOcVvA6AQAojPabAAAAcI+2bT83TXOQ6dpcN03ztOu6r5v85Yrabn5pmubZfL3caF0AACAlKvUAAADgjrZtn2ac0AtONk3oRTW03QyJzmMJPQAAciWpBwAAAP+Uc+vND13XXWz6lytqu3kyXy8/J3AdAACwE0k9AAAA+Kdck3rX28yLi203t5q9l6lf5uvlxolOAABIkaQeAAAA/FOuSb2zruuutvn7FbTdfD9fL88SuA4AAHiUtus6KwgAAABR27bPmqb5PcP1uOy67tmmf3k1W4TE5X+GvaTJXYYErTl6AACU4Cd3EQCgfKvZ4mnTNE/jC71dfXK3EuVZDxUbYQP19ubpx3v+92cbrEDCcq3S03bzW9cSegAAlERSDwCgEKvZ4llM3N3882lPSbptHdz5+4e3/vfrm/+xmi2aWwnAz03TXMV/SvgBU8sxqfe+67qPG/y9GyEBuD/sJU1OQg8AgKJovwkAkJlYXfEsbjrfJPDuJtJy9yUm+T7GPxJ9wGjatv2a2Zy5UJH2tOu6jT4nY/X2f4e/rEn9Yo4eAAClaf/fk/8bNoKeVHBnv87Xy88JXEfv4ql895B73dr4rYVN3x2U/DkyXy+3ObFelDgnpzRVfhfEzdejW0m80hJ4m7q8leT76PO+DGIVUpPpPL1fu657s+lfXs0WH+9UUZfm/Xy9PEn1NVX0DM/jXM3Xy6vH/BcKfR64j+/WCVUYy9G/R3/efU9Fn4M/UsQ+is+av1yF9ptnhQfzNz5lPBfhIe4hP3J6u9VZBX5tmmbjDQ3+UuznyGq2eDtfLzeeL1OY/xT4mqr4LojBanidx/GfpbdH29RB/POq+XOdQpLvIvxx8CdrYhVSk9v3zJctE3onhT8/Xm4zW3AitTzD8zh9fF+U+Dxwn5/vzHBmXM8qeq8xjCHj46PKnjXutZot/lXAM/Obm72AioXuHM/+T+WLADWoLZlRa/KG73sVN68gaaEabzVbnK5mi5CkWjdN8++maV5I6P3QQXxA+301W1ytZouzWP1AJmICW6xCanJL6m0c58TfuZJbUoaNjhMVOwBA82f3ppAI+mAxmvMErmFnq9niWELvjzg3zIu+ktSDgsVERk6zQPqwJ4HDPWz0k6SwuRo+s2IbtDDb6LemaZ67WzvZj0H+TYLvNG5ek7ZjsQoJyimp96nrum2qU94U/jt3onIbALjjJFby1+xgNVtk2S0kPtdnnZTsyelNnCupB2WrtbWTllbcFTavLmzwk4rQ1381W5zHirx3WnD1bj8mSNdhnc1RSJpYhaTEeXo5Jb22qdJ7VvgJ59By/SKB6wAAEhIr+E9ipVPNXmd64P28woOgd/0yXy//SmxK6kGh4gZmrS3b9m3gco/9OHsLJhOr8q7izIkX7sQowjr/J1RDqo5Ki1hFrJKonO7L267rrrb4+yW33byseIYyAPCAWOF0bJ3yqniLbTdr72b0fr5efhPHS+pBuWp/qK399XO/wzBzy9owpthi881qtvgaq/LMyJtGqIZ8F1tzSu6lofZqNbFKmnJJ6l1v8zsUP/dKrQq/3qZiEQCo03y9DC3Lf6n89mfThlPbzT98mK+X/4hzJfWgQKvZ4qlTDM3zuA5w1ysb+ozhJpnXNE2oonitXUQy9mNy77NKqenE7+ja286KVdKUy+fCWdd1Xzf5i3FDpOQk+qk5egDAJmLF0/vKFyuXNpwXle+jXH7v4JqkHpSp9pPvN6wD33OWaR9xMhETx58l85J2ENtyXkisTMJ39J+sQ0Iymqf3peu6bd47pwVXiX+4PV8EAGADpzFhUrPzePArSavZ4rTyQ6ChE8VRnAf5D5J6UJj4gaxH9J+OU/6CYlJhw+7C+4O+hWRxmN2mzWZWQmX751xakJRArPINsUpacqnS26bt5tOCW71+0XYTANhWTJQcxcRJrQ5SPWAY49ean89/mNBrJPWgSKeqQv6y50GfH9iPpfzwaLHVZmjj8buWglnaiy1IPqviHYVY5W9ilbTkkNS77Lpum8q0NwX/vp38aLMDAOB7biX2avYq0ZEU55U/L5481FpeUg/KY2PoW6WeTKYfhzERAzuLSaBQnffKKmYvnFb8Pbb6YDhilW95v6Ujh6T+xu+XuEnzYtjLmcyv8/XyY6GvDQAYQUycvKx8rZNqw6ntZvNyvl4+WIAgqQcFiTOctHv71n5cF/iecDJJGzh2EgPO32MyiHL8FmftaYvYM7HKvcQqCWjb9mkG781PXddtk8gqtW3R5Xy91DIZAHi0OJv3fcUruZ9KzBgPTP+WwKVM5f2ms6Il9aAsNoTuZ114yLmWe2wjtts8rzzgLF2YtffRZ0PvfCffz7pML4f2S9vM0jsu+JSz3xcAoDfz9TLEFp8qXtFU2nBu02K+NB/i+3AjknpQiPjha47T/Q5tyvKAvdRaDpCu+D75WHBLM/52ILHXH7HKD4lVppd6Uu/DllV6pbYX/+WhGSMAADsIB6K+VLxwk+6JrWaLNxV3QLrc9tCapB6Uw4nVHzOvhoccVH4qiA3cmp+n3WY99mJiT5vexxOr/JhYZVqpJ1W3maVXapvbT/P10ixkAKB38/Xya0zsXVe6upO14Yz7LK+n+NkJCO+3o/j+25ikHhRgNVs8VTHyoBdxneBHnsfTQfAPEnpVC4m9f5t7tjuxykbEKhNp2/ZJ4p/t77uuu9rkL8YT1iUmvq4dDAAAhhS7AdR80O7V2IdZY+xa6wH7nRJ6jaQeFMMD7masE5t4rSKHu24l9PYsTtXeSeztzLptxjpNI/UqvW0OHJ0W+l31Zr5ebpTYBADY1Xy9DAmmtxUv4NhtOGtuu3mya1t5ST3IXPyg1a5pM6dmprGhc7ONuHHr5JiEHk1M7Pl82IJYZSvWaRopz9PbtkqvxPeQtpsAwGjm62WIpz5UuuJ7Y1XOxZnrr8b4WQl6OV8vL3a9LEk9yN+xjeaN7cX1gofsTT0kmKRoucldHyX2tiJW2dyeatBJpJzUU6Un2Q0AjC/E5JeVrvvzoTtYVd52832sCN2ZpB7kz/yv7VgvNnVQcYBBtJotziX0uEfYNL+Q+N+Y797tWK/xHSZ6Xar0mubXXdsSAQDsKs45O4lzz2o09EH38MyzX+G6fpivl48+xCmpBxmLZco1fgA+xn5cN9hEOJ1kc7VS8WTai9rXge/al/h/mFhlJ2KVEbVtm3LV7TYxyJsCq/Qu5+ulOAwAmEQ8WFRrF43B2nDGvZYa225e9vV+ktSDvHnI3Y11Yxuvh247QHoqbwXB5kLiX1u4H/Oduxvvq/GkmtTbpkrvaaEbI34PAIBJxblnv1R6F3pvw1nxXkuo+DyKFaCPJqkHmYoP76m2CkrdYVw/2NS5+VnVOTcDjA395vPhfmKVR3kuVhlNqlWR21bplebtfL38WODrAgAyM18vz8KBq0rvW99tOGvca+k1oddI6kHWnHx/HOvHNvZG6CdOImLbu+fuB1tQ1Xk/37WPY/3GkWJSftsqvdJaRV97/wMAiTmN7RNr88c8+T5ec6z6q3Gv5aTvGdGSepChmFjQDvBxjiVo2NKBjftquM9s68D8zW+JVXohVhlY27ZP4vd7arb5Hirxs+e0z5PMAACPFWOT43j4qDaHjx07UXHbzZexhWuvJPUgT6fawj3anjkd7OC5jfuyrWaLMLR4v/Z1YCen2iV+Q6zyeHsVD+YfS4pVep+6rtuo7WShVXqf5uulwzUAQHLm6+VVwq3bh/bmkc+7FxU+H74fKq6V1IM82eDph3VkF6/7HhRMUiRt2dWe9883fMf2wwGkYaW4KbPN50iJv2fe8wBAsmIbxZcV3qG9XSvtYpVfbbPWP8zXy8FidUk9yIwqkl7tx/WEbYX5eime7ucRYrLW5yuP8UK1nlilZ2KVYaX2Xb5Nld6TAhNgb/ueNwIA0LdYffW+woXdug1nfD6u7fDr5dCH7yT1ID82dvplPdnFXkzsmXVUFp8H9EG1nt+lvlnP4aSW1Nvm9HNpLW6vfX4CALmIVViXFd6wbdtwnlfWdjPEtEdDz4eW1IOMrGaLowrLlYd2GNcVtnVQ6ZDfIsUE7fPa1yEKQeinDf/wT1VX64lVBnGoOrx/bds+Sayi9EvXdRvFFYVW6Z0OvfkBANCz8OzzpbJF3bgNZ4VtN0dJ6AU/Df0DgF45qT2MsK4btTqCO56vZos38/XSyfL81Tgn8Tp+9n2O/7yKg7+3FhM5T2PVy1FMetfstOK5UGKVYZxa296llijdJpYorUrvMraxAgDIRkjexDEeHyurRvujDed8vTz73l+IhxJ/G/eyJncyVit5lXqQiXjq/4X7NQgzkHiM1zGII2813cMPTdP873y9fDJfL49DUnq+Xn7cNaHX/PkwE/798/l6GQL7ELzP4vDwD/1eejaqTL6IVQYlVulfSp0awiGLiy3+fmmfMbUeggAAMheTODXGMr890E2ktgNbL+fr5Tbx/KNI6kE+nM4elvXlMc61RsteDa03Q7//f8VE3qDBZjixGJN8xzHB92vctK7F3mq2qPF7xXfpsKxvv1JK6p11XbdRm5742ZJS29DH+hAOhuT9EgCAmsWOA28rXIJ7E3eho1Vl3Xvej911QlIPMlDo3IzUnMZ1hl3sxcSe91CGKpmrGYLMZ2O1grgtJvjexPacNT3oVFXBK1YZhfXtV0qHcbbZBCit5bf3NQCQvdC1psJONQcxgfeXeOD99bSXNapwQG30w5dm6kEejivrzTyFvbjO5nmwq4P4/tGKMz+lV1leThFk3hWHRYcDFBexzVzp32th5uaTMYZkJ0KsMrw/KkDNHnu8tm2fJPR+fd913Ubtj2O775Kq9N4+pvUzWfg0Xy9rODwFVGK+XrbuNT9wEufr1VSlFkbSXNw6QFzTs8rlVN1UVOpBHko7kZsq68xjPb97SokslJ7US6oKIrZZO6qkHWdNSX6ffeOwzv3ItUqvpKq2a+9nAKAk8UDnSWWjJ5qbeHY1W5xVlNAM9/hoqkO8knqQuNgWrqQTuSnbr6QNH8N6HU/Sk4+nBd+r6xRnFcVTfDUk9qr4LBCrjEqs0o9U1vCy67qNPqNjK6PD4S9pNGcVVTIDAJWIz7q1zcI+iB15XiVwLWOYNKHXSOpBFpxgHZf1pg/ncfONPJR8r0afobep+LBT+mduLckX353jMoPs8VL53D/b4u+WdN+/bPnaAQCyMV8vQ4Lrl8ru2PMErmEsJ7fajU5CUg8StpotnhZ2IjcHh3Hd4TH2YmLviVXMQslzwJJ+D87Xy7O4uVuqvdIT/GKVSTwXqzxaCr+X113XbdR6M8YTL4a/pNG8UaUHAJQsPuu+d5OL8zImbSclqQdpc/J9GtadPhxUNiCYNB1kkFwu/TO39Kpd35nTsO47atv2SSLtYmudpfdlvl6KjwCAGoQY7tKdLsb7VOJYST1IVNyENZdrGscqrOhJqKaw8ZqwStqkpv5dMvkpt4EV+x4rsHooJ2KV3eXYerOkuSziIgCgCrEzwXEFs+Rr8GG+XiYTk0vqQbpOC28Jl7I982ro0evVbCFBn64aNsXPUt78jw86nxK4lKGUnDj2XTkdscruUvid/NB13dUmf3E1W5wkUlnYB1V6AEBV5uvlVUWz1kt1mdohu58SuAbgfiWdyM3RiZPE9CjM1zuaepAu1Qqb/x/jezDVGUbhM7fUGWElz40Sq0xLrLKbFJJ62yS2VOkBAGQs7AWtZouXTdO8cx+zE6osk9tLkdSDBBV2IjdX++E+OE1MT/ZuJfZK3uAnXQcxsXccTwomJV5TctfF94lVkiBW2c3UBwi+dF23Udvh2CL6cPhLGoUqPQCgWiEOCntCxhdkJcmEXqP9JiQrxxO5JbZOU4FAnw62PJnPOGpKsob3YDghqGUffcj9ffQl/smdWGV7UyfJtokFSvq8VqUHAFQtzmS7rH0dMnKcasctST1ITDy1keOJ3LMCE3uH8X5AX56vZgubWgmpsCVqqBr9bTVbfPb5xq7ie+cg8wW8KOSghVhlC23bZtN6M85CLWUmryo9AIA/HRVyuLB0L+fr5cdUX6OkHqQnxxPX1/P1spTNsbucgKdvr0MLRKvKxEJC5j+r2eIqtO+Lm8ewqRK+G88LilvEKpubuvXmp67rNm01fBwPYpTAgSYAgD8PFn+Ncd619UjW29QPpEnqQUJWs8XTTHsrX9z5Z0lexPsCfTqPc3JIQ83B9H4c1h2Se+cSzjwk41jltlA19DnOciyh/Y1YZXNTf/fW2HpTlR4AwC2xY5CxGGl6P18vk783PyVwDWN5UnBrGqfry5HrSevQevOP0yar2eJDaDE4/SX16sQJY3q2FxN7SQ7crdDnTNse92kvJmpCcuA6HtYIrSYuvEe5o4SHz7Nb/zskG36b8Fr6IlbZzJTPg9dd123aevNZAS1ub3hfAgDcEQ49xZjvlbVJxmUuz7s1JfX+aDOVwHXAvWLrsxw3yr7cmUl1XmBS73Q1W5zZ2KZnB/H3RWXU9CT1vvVXgi9U8a1mi8uY4Pvjj8/CesVYpYRWj7c7C5SS1BOrbGbKisZtOlqo0gMAKFyoCIsdN0rbR81RmHOYzcF77TchHbnOzfhmgyLO1iutld2exAsDeb6aLZxgn96m841qdRBPD/67aZr1arb4HJIHcRafln91KWHG12Vsu/mH+ND2YdpL6oVYZTP7E/7ssw3+zo1S7uU2rxkAoEYnhYwEyFnYxz7O6YCkpB6kI9eN/ftO35Y4W0/ihaG8Nsdsch8rf/3buknyhVl8/13NFqH18kVIUIeWsrGaizKV8F1YctwiVvmBtm2nbL35peu6zxv8vVARe1JA8ryJmyOq9AAAfiAmkk4qn/U/teM7XeiSJ6kHCYjzHqc8Obyru603b5R4Kne/4LmcTO+mlzoTiJ9jAujd7cV2Ia9jq/NQzXcl0VeWePggx1jlrn8k8GJ7wBI+A8QqPzZlZfE2ya1iqvS0gwUAeFjckyhhzEGOXs7Xy+wOekvqQRpyPVl978n2+GX0ZfzLGZwT8AxlLyb2JD6mU2KF8ZT2JfqKU8KMr29ab96hWq98ySf1CpupovUmAMCG4jijX6zXqN7mOv/5pwSuAaoWH94PM12DHz2sX8T2bCU5DPfrBxuC8BgHcdNPK85phJNZL2p84SPav5PsC9+BX+Lah8Mgn3M8IVeDzGOV2x6KW0r4DBCrfN9UVYyXXddtej9KiQHeq9IDoERm4m/kXCy6m/l6eRa7ONmbGF6IV7M9uCqpB9PLNSD40Wn3Jm6clZbUa+L9UhLPUJ6Hh4T5eulBYXwX8XOrhDlGOdmPDyx/PLSsZosmDgn/fJPsy623faFK+Uz6bjVeOBkbk8wltBgVq9xvqkq9bSrWSqiIbVSMAlCw127ug8JznKTe7kI8+Cwe/GYYl7nH3dpvwoRi67FcT+T+sDw5Jvwux7uc0RxrGcfAXsfZVYwoVhRowZmGg5jke9c0ze+r2eLrarb4eNO2s/bFGVv8zivhpOiHDSqHSvkMEKvcb6qE7Ubvq3gqu4Sk8ien8wEAdhOfWY7N/R9MOMh5lHtXCUk9mNZpxlUhm2xQZNmX+AF7BZ2iJl3ncXOPcZn/k6a92Prxj/l8q9mik+QbVSnfeTXFLWKVO9q2neqz4kPXdZtuGJRSXem7FADgEeIBKc+6/QuJ0uMS2sRL6s6yjccAAB9RSURBVMG0cn14f6j15o0Sk3qNllaMYC8m9lRajCi2efxUzQvO231JvlPJ8EGU8p33YFIvfgaU0mVArPKtqVpvblP9WUKV/pfQyjaB6wAAyFp8NnnpLvbquJTxHpJ6MJHVbHGScYudjZJ18eTDh+EvZ3T78f7BkA4KToynzBygPIUk32+32nWGpLgWhI+Ueaxy2yatN2+U8rkrVvlW0km92Ha7hN81VXoAAD2Zr5fh2eS99ezFy/l6+bGA1/GHnxK4BqhVzhst25zADX/3+YDXMpUTCRdG8DxUH83XS5tkIwlB3mq2eF/IDLFa7cX798c9XM0WH+J30UUJbTZGVlPrzdt/97cBr2VMYpW/TVHFu03rzVJm6Xq/cddhqKiveFV+LmkDEYDxzdfLk9iR5sDy7+xtTJAWQ6UeTCDOADrMdO0/bDn8/qLQ4a6HZjkxkt+810b3xlDqooSDJe+aplmvZouLUL2kgu9h8XOnhAfH620e4GKMU0obXrHK36ao1Kut9eZ7BycAAAYRYvovlnYnIUYtbt64pB5Mo5YqvZsWnKXO1tDWirGERMRUrcOqEzf1teEs0+0E33lsecf9qpmld4+STnGKVf40RYJ6m9abe8NfzuB0FQAAGEDcWz12+HhrlwV1n/mGpB6MLG7M59zWbZfNsVKTei8kWhjJXkzsqS4aSWx5WuJMUP4Wvov/vZotrlazxRu/X38rIFa5rfa4pfpYpW3bKV5/ba03L+fr5ecErgMAoEgx1ioyQTWQUNl4VGonCUk9GF/OH8AfdvkwnK+XFwWXiTsBz1gOnIIf3YkWF1XYb5rm9a3qPYc1ynlYvI4xyFZirFPSQPraYxWtN4cnPgEAGFgcK/DWOj8oVDQel9waXlIPRhSrAKppvdnjv5uyU9UdjChUXDiZNRItLqoUqtP+u5otPtY6i6yAWOU2ccufao9Vnk3wM2tqvXldcFcOAICkxPlwpcwAH8px6V0kJPVgXLk/uD/mgb2k+TS37RVywrpUJSZjfqs12TCFGAj6Ha/PYdM0/6k0uVfKfK/mMXFLrPAr5Tuk9lhl7Eq9y8pab16UfAoaACBBx3FeHP/0cr5efix9XST1YFxvMl7vnVpv3ogb46W2scv5vpau1M34Cy0CxxMDwpe1vF6+cZPcq6ktZynfaTu13ryjpOqjmmOVsSv1tjnIVkJSr9SDewAASYr7sye6Cv3D29iitHiSejCS2F5nP+P17uNDsdR5G/vx/pKYmEwuMRmzFxN7Wr+OJAaGEnv1Cm05P69mi6ITIwXEKreJW761X3GV99gJ+Zpab36p4SQ0AEBq4n5X7bOzb3sfW5NWQVIPxpPzB0sfp92bwudtmHOWqJiMeV/gSzsoOFGepPhe+pfTcNUKm++vV7PFVcHJkZK+yx6d1Cuwy0Ct1XpjJqq/dF13teHfLeFzRJUeAMBE4l7tW+v/xzNbVfuyknowgtiy6zDjte4lGTdfL68KHuZ6qB1iuubr5Umh/cZfrGYLCeURxU3+I/3rq7YfW3IWlVQvIFa57UuPg9FLSlpUF6u0bTt2681tYmatNwEA2FmM7VXr/fmMXlUHNUk9GEfuJ6P7rLAr+eHfbL20HRVaYfVbxS3VJnErsfehwpfP316tZovPBSVJSvoOE7d8X22xythtqjdtvfmsgFa3n+KBPQAARhbHsVwU0M69L2cxxq6CpB4MLH7Ivsh4nftqvXmj5BacL8w4S1ccJFxq8utCpei4wvtpvl6Gk2C/aMdZtYM4ay/rz5YCYpW7ekvExaRFSZW5x5XFKmM+2F93XbfpfDmtNwEAeIzz+DzKn/bi3lgVzzqSejC83Fvj9ZqEi4mVkqtbtEJMWKywelngS6sqeEnJfL08i5vGpbYW5mF7sR1nzm1PSvru6rP15o2Skhd7lcUqY34vbprQawppk1TyQT0AgGTFURDP3aF/2K8lRpXUg+Hl/tA+xEZWySd79bJO3Hy9DO+/9wW+tHBCq6gZX7kIlTzz9fJI1V713mWc2Cvpu0vc8rCaYpUxK+I2bb35pIBT1R/iQT0AAEYUnzlfWfPvOixt/v19JPVgQPGDNud5GeG0+zanjjcS23mWuvG9n3m1RhXm6+VJYe3UboQWsKpFJxKr9kIb1LdVLgBNjom9AmKVu3pPwBXYZUCsMoxNY+YShvhrvQkAMLI4M+6ddX/Qq9KfdyT1YFhab07z356apEoejgpNLv+W+3yvnMVZe+Ez4H8KrQjlYbkl9kr6zrqMM/CGUFrcUktS73Ckn3PZdd2m773cv6P7nrcNAMADVrPF0y3bvdfuLCZBiySpBwOJm+q5t9YZ8hRuyaXQB5Iq6YuVF6Xep4sY8DGR2JLzRHKvWiGxl3w1TiGxym2DxS2xdXNJB0EOS49V2rZNdZ5e7pV6EnoAACOK7dsv4nxsNrMX98bGfCYYjaQeDCf3E9Ch9ebnof7j8b/9Zaj/fgK0tcpAfB++LPCl7dl0S8M9yT0z9+pxnsHJwNK+q4b+3FOtl5cxf/82SurFz4TcN2PEFwAA4zov7DDmWPZLjV0l9WAAsULmReZrO8aHXsmbAi9USuUhVl+UWEkl4EvIreRe+Fz4pfBDDfxpL+W5U4XEKrcN2XrzRmlxS+mxymincruu2/S9kXuVntabAAAjWs0WodPZc2u+s8O4hkWR1INhlDCfZoyNyJJbcDZm6+UjJlsua18Hhhdn7p3N18uwkf6z6r3ihXbMbxJ9kaV9Rw0et8RkRmkJ+ZKr9caq1Pu0xd/NveWphB4AwEjirPZX1vvRXmU29/5BknrQs9irN/cPisshW2/eiCfqS06knJTau7lQR5IrjGm+Xn68Vb0X2sB+cAOK9Dq1NpyFxCp3jVUVWVpS47TgWGWs17Vp681wPYfDX86gJPUAAEYQnyHfWevenGUwHmNjknrQv+MCZmWM2S4s2dZkPdgroM1SNUIFVQEn6MlQrN47n6+X4fNiJsFXpNQq00uIVW77ED/Dx1Ba3FJyrDLWQ/tGSb1CYoxNXysAADuKLfLFXf0Kzz0XpRxolNSD/qXaZmsbY57CLf3Ebwnvh2rECtWXta8D07md4Juvl23TNP8bW3SawZe30Mc/pQ390r6bRosl4vdEaV0GxCqP0HVdLUm9MZPnAABVikmni8IOYaZiv5R9aEk96NFqtjiOHxA5u4xtMUcRf1bJFSn78X1BJkJCJSZRYHJhhldo0Rln8P2raZpfVPFlK4nESSGxyl1jP5iVVq23n1jSuS9jtLo0Tw8AgD6FZ40DKzqYcOA2tU46W5PUg36dFrCeU2xUlb5JUML7oipxzlnJ8x7JUKgQmq+XZ7eq+H5umubXLTeVmU4q1XqlfSdNUT1UYtyiWm8328zTy31zRgsoAIABxWTTc2s8uFer2SLrGfOSetCT2O849+H3zUQbVeFnXk/wc8dyGN8f5OWo8PclmZuvlx/n6+Wb+Xp5FJN8tyv5tOtM06QPDgXFKreNHrfELgOlJdOLilXath0rgV5L681RO3kAANQmJpleufGjOVvNFmPN4O6dpB70p4QTzpM8sMcT9mbrkZT4viyxHRmFulPJFzbn/yfO5Lup5pOknt7U7ZhL/C6aKn4orQVnI1bZXkXz9LTeBAAYSEwuvbO+owozCy9iR43s/FTHPYJhxQ+AFwUs85Q9hS8KWcPvebGaLU4naBHGI4QkyWq2eCm4IkfxkMbV7c3YWInzLP45iv80gHs8e2GmXZiVOPYPLihWue39hN+rFwV+NxyH90khscoYp27N0wOA+vzqnj9Idf8W4jO6NufT2I9xbnaxek1JvU+hPVUC19G71WzxscBWSrkpZT7NZA/sYYNzNVtcF765fOoUfH7m6+V5nINVctKZSkj0JeF4ou/bEue7Thm3fF3NFu8L+27YKyhWGePEbS3z9K7DIacErgMAJhfGH7gL9CXGiReevycVxhCEjkdZPS9rvwn9yHq4ZvQhgZPZJbayuq2E90mV5utluHeXta8DZQqJvnCw4tZ8vif3tO40o68/Ux0yK+076HqKisc7SqxeKuV9MsZ8wE0TXar0AAC4z3nmh79K8SrONMyGpB48Uvyl3y9gHVN4YC89qbef25cE3zgyk4xa3JPoM6OvP/tj9+0vKFa5bfK4JSYVS/s9KCVWGSOpt2mbpGwH8EfaQQEA9CxUhzVN89y6JuMszjbMgqQePJ7Wmz2JrX1KrwYpsf1ZFWIla5FtnGETP6joC3Mn36pm3crYDwtabw5HtV6dvnRdt2mHC5V6AAD8JR6ie2VFkhJaoF6MfQB3V5J68AhxzlYJZdIptN68cZbGZQzmIL5vyFBMPL907+BPMdF3HvrPz9fLkKiaxWo+Sb4fGy2pV1CsclsKrTdvlBi3HBYQqww9b3yb6rWcZ59fJvSMAACQvVgN9s6dTNJ+LgfafkrgGiBnpZxkTukDK1zLbwlcx5BOtDLKV0hgxM3OF7WvBdwVN38vbr5XVrPF01ilcqy1yDfGPP2nSm9A4bDHarb4UmB7U7HKj200T6+A5Kj3AI91WXmnkk1nbwJQgfh8LL5KWzjgeBYOLqd8lZJ6sKP4QVzKpv7T1WzxJoHruHEdy55L9SKsd6hwKfg1Fm2+Xp7E01UGGsMPxM+5MC/1PLaxOI7JgpwrV/owSqVejFWKTKYmFrd8LTCpl22s0rbtGL9fm27Um6dH7b7O10vvIwCqF5+HLwrf7yzFq9Vs8Tkc6k/19Ujqwe5KOnH4OoFrqM2p+XrZC6fvrwRksJlYxXeT4AvJpjcxyVfj79BYlXqlfs+olB5HrrHK4L9fXddtmqSQ1AMAoInPwg6G5+MsJvaSrLo3Uw92EE9XlNJ6k2mc5DJ8lfvFBIX5iLCDOIsvfI+G5N57a9g/sQo9EKvc79MWfzfnOME8PQCAHoR2jsZRZCccPr5I9XlIUg92U2tlAf3Zi+8jMhZP7Lx0D2E3YcM4Jvf+Fefu0J8TsQqPlGusMnQibdN5ek8yb8uqSg8A4JFWs0V4LntlHbO0n9I899sk9WA3Kc1xIV/eRwWIPbZVGsEjxAT5kcRer7R4pg9ilX8yTw8AgAetZosQD76zUlk7jJWWSTFTD7a0mi2OMz91Szr2w/tpvl4meeqDzYVKoxis6Y+egdVsUfJGZah8y7IKOFTtrWaLo7iR7HfpEcQq9CjHWOXpwP/9TZN6ubfoltQDANhRnCMvnirDqzhf7zyVVyOpB9tz8p0+naZays3WwubdlXZ3WTgs+cWFBHOqw5wfEhN7ISH137SvNHliFfqUW6wyaFKv67oaKvXM0wMA2FFsw35hf6goZzGxl8Rei/absIV4yqLozWBGdxjfV2Qubn7lfiqfMpzk/Crm6+WVlra7i1XDYhX6JFb52zYtgnNO6jlVDgCwu3PdZ4oTErQXMWE7OUk92I65IgzB+6oQ8cTOy9rXIQOlz207SSXQfITSK5iH3DBXpccQcopVhkymbXQyN34G59wCN8tqbwCAqcX5a8/diCLtp7JXIakHG4oP5y+sFwN4UcAGPFHssa3KKG1Xhb++vQISO9q+7UCswoByilWGbHO06fdHzlV6jUo9AIDtrWaL0DXnVUVLd53ANYztMCZuJyWpB5tz8p0heX8VZL5enlRQDZazGioQTrXLS9pQ70HfJQzJ+2vzZFfOSb3r2AYZAIANxTEI7ypbr7D39WsC1zG2VzGBOxlJPdhc1jOKSJ73V3mOKj21lIMaknpJ9XvfQe5VLg8ZasNc0oUhJR+rtG079GfHpt8f5ukBAFQiHqitLYb6MF8vL+brZWjT/ymB6xnbWUzkTkJSDzYQs+85z8UgfftTn/KgX/P18mtM7JGeWoLtMJh78rYQOzrO8qo3FOdv9ip+hwzZdhByiFWGPMhw3XXdpq2Bc66UNk8PAGBD8SDtRWXPYl/uHPg7qfBQ+6QHqSX1YDNOvjMG77PCxI37l7WvQ2piwrWW9qhhDtZ5AtexsZg0OMzkcncx1HvPdwhjqPl9tk2yK+fPMJV6AACbO48HamtyEvdV/hBbt9f4nLAfE7qjk9SDB6xmi6MKP5yZxkF8v1GQ+XoZArz37mlyskp0PVJI7H3OYcZebF+Ra3XhpnrfMBerMKLUY5UhW+BslNQrYJ6pSj0AgA2sZovw7Pq8srX6db5e/uOZNu59fZjmkiZ1GN8Ho5LUg4dpiciYvN8KNF8vTyqqDMvFJKepJhQSPp9Tbp23mi2OY8Kr9LYlQ1TBqNJjTCnHKkO2v9l0FmbOSb0vt09dAwBwv/hs/aqy5bmMM/S+p8Y2nMGrsfdafhrzh0Fu4knbF24cIwoVNW9i6TplOYobgmZeJSD8jq1mi8vKqpvCe+9dDDbf3He6bgqxB/1ZRd+3va57jFVqOx3KtGqNVTatYMu564IqPSBX/1nNFm7e7n5O5dkgdavZoqt9DcYwXy/blK8vdph5l8CljOn6ocN94XBYPKz7n+Je/cPOQoekIebn30elHvyYk+9MwfuuQPHku/aqaSm9zeP3HMaNj49TVu6FZFSc93dVUULvwwBVML4zmEKq77shq+Q2fUDPuVJPUg8A4Afiocoak+BvNklYxQMCb8e5pKSEQ9QX8dDy4CT14DviL6FWiEzhZKwvAcYVA6CXlj0ZF5W2hrhxGCv3wmm683iiblDhRONqtjgNJ9iapvlvTObVVL3aa9tXsQoTSjVWGSyh1nXdpgn5nJN6qjQAAL4jxr8XFXZg+jRfL7c5FP2m0hE0+2ONetF+E77vWJs8JrIX33/nbkB5wvDg2Kqhtt7ryYmtIUJg+rrypdiLybUXsW3Rp1itEf5c7dKKJz7sPIub209jleqzyr9XrwcI8E/EKkyktljl0xZ/93DA6xia9u8AAN93XtkIjyY+x251ADjutYRn1d+Hu6xkHYZ9pvl6OWhnE0k9+L4fDf6Eob2R1CtX+HKPib2cN/5KcRbbyEmM/O3w9nvz1nySTTa1a0/c/ciF1psUpqZYZaPf3dw7LZjpDABwv3gguMZZ5ie7PMeGTlWr2eKXpml+G+aykvYqztcb7FlJ+024R2xBtm9tmND+GK3wmFS4v1/cgmnF4LTW2XrbOtzgj4Te9/V6WEisQgJSjFWGOiyz6ay5ZwP9/DFsU40IAFCNWHVWY7el9/P1cuduM7FlZ60x5lk8zD8IST24n5PvpMD7sGAxmXRc+Uy3JMzXyzcSrAzs/QAVML4jSEEt78NNf39znqe3aeISAKAaMTHzrsI7/qWnWP+k0n2vcOD5YqhOHpJ6cIeWeCQk9GHOeXOIB4R2BDbmk3FS+wIwqL6r9MQqpKKWWKWGpJ7WmwAAt8Q4d+sZ84XYqe3mXfFwa637LfsDzNX/g6Qe/JMNdlJitmPhYo/tt7Wvw9Tm62UI1D/UvQoMRJUepUsiVmnbdsh5djW031SpBwAQxQqri0pHTPwa90h6EVt41rrfchjnMfZKUg9uiR/YL6wJCXkxVKk26Zivl6dm2SThRBtOenbddwJOrEKCUolVBkuodV236SnlnGM2lXoAAH8LB7APKlyPyziipG8177e8inMZeyOpB99y8p0UeV/W4VhCaVq35hxCX077aFlyh+8EUlTy+/Jyi7+bbVvcASqKAQCyFCurnld4966HapUZn4trHntyFsdo9EJSD75lo4wUeV9W4FZCqcYBwsmIcw5/qX0d6MWH2F63b74TSFHJD+h9J+ZTpFsAAMCfCb0Q176qdC3exD2RQcSWnrWOnwltXC/66nAiqQdR/NCusU8y6dvru0ybNMXgyYb9xObrZTiV977qReCxBjnhKFYhYfsFxyobbWysZouj4S9lMDUkLgEAfihWUr2rdJU+xb2QQcXxM9t0wijJfpzT+GiSevA3G+mkzPuzErGyp9aTS8mYr5cnFQeaPN7RAG03G98FJG7q9+dQSbUaEl6DncgGAMjBarZ42jTNx0pv1vXIo0hqLlw4jO1dH0VSD/4+WVvj8FPycZD5CXC2EE8uaYU1vSOJPXbwcoiWJWIVMlBqrLLprLneZmRMwDw9AKBasSXiRcVdUU4GOpR6L2NPmleP7XLyU3/XAlkr+eT7l/l6+TSB6xhUPFHz34JfYhPfp7WeGqrRcTw5v1/7QkwlBLVxg/qjZAobejvQHL1GrJK/SmKVkwJjlU0TXr3Mx5iIpB4AULPzip/5wzNsLy0htxFafa5mi7DvdTj2z05EeP2fdz0QrFKP6sUNlucFr8PoH8xTmK+XVxVU1DyP71cqEE9JHcc2CEwk3gcVe2zifayy7Z1YpQyVxCovCoxVNj21nHOlnvabAECVYivEkp+1fuRL0zRvJvz5JxXveYWq0ItYJbo1ST0ofz7NUBUDKRp8oGsCzFOqSDyx455PTGKPDbyPcxiHIlYph1hlOINUynVdt2nCK9tKvTHbLQEApCK2QHxV8Q05njIOjIcea56vt7/rAVdJPaoWs+Elf3h8GWKuT8JqOOl/suspDvIUW/m9dfumFQLd+XoZqjDe17wO3GvQhJ5YpThileFMXSmXa3zmwAoAUJ3VbBFix3cV3/lfU3gOi60/P0x9HRM6jNWiW5HUo3YnhQ9BraKd1Y14uqT0L4K9yk+xVCm29PtU+zqkICZvfq19HfjLy4Er9BqxSlkqilWOE7iOPmzz3ZvrHBZVegBAVWK7+NLmQG/jcr5eTtl2866T2Aq0Vq9i1ejGJPWoXentrGpo8XRXDS28tGOs03HlQU4yYvD7v+YdVi3c+59jJe3QxCrlqSGRmdImAT92ZX0AgFrEjhIXhR+c/JHr1A7gxYOPtRcwnMXq0Y1I6lGt1WxxHHvXluoy9iauSizbLn2jfT++f6lIDHKOJZLSED9rnmlbVqVwz5/N18vBT3aKVcoUk8FilTxsVMW2zQN4giT1AICanGfcYaEPb1J8BovP1zV3RQpJ5otNxxhI6lGz0k++11Cx9j01nIBXrVeh2O/cvU9ECITjnD3tOOvxNtzzER+CxCrlEqvkYdM5IznPO9Z+EwCowmq2CN0knld8tz/M18tkO6XErkg1H5ze3/Q5UVKPKsXTtIeFv/aqZtTcUcMm4WHmp8LZUazweGv90hEDz3+p2ival9huc7QkhVileLXEKk9H/Hniot1smrgEAMhWnFn2uuI7eJ1Ji8uTyjtUhWeoBxOvknrUqvST71W2s7oRS7ZrmD2mYqtSMbHwqfZ1SEmoorxVtadFall+Havd5h1ilYJVFKuMOVtviLkom1axHQ3wswEA6EE8MFnjLO/bTuJYl6TFDlW1z+d+FZPQ3yWpR3Vib9oXhb/umttZ3ajh9P+LTXstU6TjSjaEsxKr9sIDw/va16IAIXH+P+Gejv3wI1aphlglfTVUsanUAwCKFWPRjwMdAMtFGCORzbNHbBFa+0H2sx91aJPUo0Y1VDfV3M7qRi0ncFTrVSomGY5VhaUnztoLp6p+Fohm6VNstXk0YSVZDZ/tknr1rIFYJWE5nNgGANiFhN4fvmRa+Vb7fld4z15874CkpB41Kn1j4UPN7axuxDWoYb6VjbKKxbYE3gOJCu31QmJIci8bt5N5Y7favKuGWKX6REL8DK8hVslhdsdjmecHAJCecOD/oPL7cpzjs1e85hqeI35k/3uFO5J6VCX2oy39dIYqvb/VcAJ+76E+y5Rtvl6G9/lbtzldknvJ+5BQMk+sUp8aYpX9XGOVrus2/UzItcWo7yQAoEir2eJNBSMNHvJrPEiYpdgytPbRJoer2eIf3egk9aiN1pt1qWUtVGpVbr5entqYS9+t5N6/BKaTu47J8DAz7ziFZN4tYpW6iFUAAKBH8UDZ68rX9FOc+Z+709hCtGav7h6SlNSjGqvZ4qiCkmvtrG6JLTg/JHNBwzmI72/qdizQyUM4KRdn7s2apvnFfRtV+E54OV8vn4RkeGrtqsUq9RGr9KNt26nbX+ZaqQcAUJTVbPEstt2s2XUprSu14fzLWXxv/0FSj5o4+V4nJ+CpQgx0ah8knJVwz+br5dl8vXwaW3O+d/8GcRmTpzdVeSm3OxSr1KmWNRnyYXyIpNo2n8e5JuNTqlIGAHiU1WzxJMY3pY8zeEhyB1gfI3bW+TXfV9CL8J6+iO9xST3qsJotwobp88Jf7LWNsntdVLJJ/jy+z6lY7JUuwZuh2JrzJFSQNU3zvxJ8j/ZHRV5M5D2LydOkH2pqiVUST6pOpZb47UVmsUq280cAAGojofeXDyU+c8VWopcJXMqU9m+eHSX1qEUVJ9+1s/qnuCaq9ahGDN7euuP5CsOgbyX4fo73s/bg9SGXcZ1+nq+X7U1FXmanE1XpVSrGKrXM2RSrAAAwhLMKRhk8pJi2m99x4vBzc7iaLc5+ig9WNcwAKDnZ4R4+7LyCjaRiyqoH8Ca+B0r3mN+RWj5HqhBmha1mi1I/86o6vBDbTPzRHi1WuBzd+rM//RVO4jpW0PyxNnGNSiBWqZtY5XE+x0MQU11r3z97LH4np+MZnrHk+vnEuB5bnT7E9zBs67ySePpHrkou+AjdqeKc7ur3L9uu6xK4DAAAthHbi4SA9tmtf5bWauRL3CT4609JswEAAAAAtiGpBwBQiJjoexb/PL31z5Sr+i7jSf7Pt/55FWdEAgAAABBJ6gEAVCC273waW1U8i6/45v/X3Pq/H5MAvLyn1dbt9pg3ibumoLaZAAAAAMNrmub/A3V5BHtK4jQZAAAAAElFTkSuQmCC";

function Logo({ height = 32 }) {
  return (
    <img
      src={HF_LOGO_SRC}
      alt="HANSA-FLEX"
      style={{ height, width: "auto", display: "block" }}
    />
  );
}

function LanguageSwitcher({ lang, onChange, label }) {
  return (
    <div className="flex items-center gap-0.5 rounded p-0.5" style={{ border: "1px solid var(--hf-border)" }} role="group" aria-label={label}>
      {LANGUAGES.map((l) => (
        <button
          key={l.code}
          onClick={() => onChange(l.code)}
          aria-label={l.label}
          aria-pressed={lang === l.code}
          title={l.label}
          className="rounded flex items-center justify-center"
          style={{
            width: 34,
            height: 30,
            fontSize: "16px",
            lineHeight: 1,
            background: lang === l.code ? "var(--hf-bg-light)" : "transparent",
            border: lang === l.code ? "1px solid var(--hf-red)" : "1px solid transparent",
          }}
        >
          <span>{l.flag}</span>
        </button>
      ))}
    </div>
  );
}

function EmptyState({ title, message }) {
  return (
    <div
      className="flex flex-col items-center justify-center text-center py-12 px-4 rounded"
      style={{ background: "var(--hf-bg-light)", border: "1px dashed var(--hf-border)" }}
    >
      <Search size={28} style={{ color: "#999999" }} />
      <p className="hf-subhead mt-3" style={{ color: "var(--hf-darkgrey)" }}>
        {title}
      </p>
      <p className="hf-caption mt-1">{message}</p>
    </div>
  );
}

function ResultCard({ item, isFav, onToggleFav, onOpenDetail, onEdit, t }) {
  return (
    <div className="hf-card p-3 sm:p-4 flex flex-col gap-2">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <div className="text-[11px] font-semibold uppercase tracking-wide truncate" style={{ color: "var(--hf-grey-dark)" }}>
            {item.standard}
          </div>
          <div className="font-bold text-sm sm:text-base leading-snug" style={{ color: "#000000" }}>
            {item.series}
          </div>
          <div className="hf-caption">{variantShortLabel(item.uitvoering, t)}</div>
        </div>
        <div className="flex-shrink-0 flex items-center gap-0.5">
          {hasNote(item) && (
            <span title={t.noteIndicatorAria} aria-label={t.noteIndicatorAria} className="p-1">
              <AlertTriangle size={17} color="var(--hf-red)" />
            </span>
          )}
          <button
            onClick={() => onEdit(item)}
            aria-label={t.btnEdit}
            title={t.btnEdit}
            className="p-1"
          >
            <Pencil size={18} color="#34393b" />
          </button>
          <button
            onClick={() => onToggleFav(item.id)}
            aria-label={isFav ? t.favAriaRemove : t.favAriaSave}
            className="p-1"
          >
            <Star size={20} fill={isFav ? "#d7102d" : "none"} color={isFav ? "#d7102d" : "#999999"} />
          </button>
        </div>
      </div>

      <div className="flex flex-wrap gap-1">
        <span className="hf-chip px-2 py-0.5 rounded">DN {item.dn}</span>
        {item.mfrCodes.map((c) => (
          <span key={c} className="hf-chip px-2 py-0.5 rounded">
            {t.mfrChipPrefix} {c}
          </span>
        ))}
        {fittingTypeLabel(item, t) && (
          <span className="hf-chip px-2 py-0.5 rounded font-semibold" style={{ color: "var(--hf-red)", borderColor: "var(--hf-red)" }}>
            {fittingTypeLabel(item, t)}
          </span>
        )}
        {item.material === "rvs" && (
          <span className="hf-chip px-2 py-0.5 rounded font-semibold" style={{ color: "var(--hf-red)", borderColor: "var(--hf-red)" }}>
            {t.materialChipRvs}
          </span>
        )}
        {item.assemblyType === "stage" && (
          <span className="hf-chip px-2 py-0.5 rounded font-semibold">
            {t.stageChipLabel}
          </span>
        )}
      </div>

      <div className="text-sm">
        <span style={{ color: "var(--hf-grey-dark)" }}>{t.labelHose} </span>
        <span className="font-semibold">{item.ferrule}</span>
      </div>

      <div className="flex items-end justify-between mt-1 pt-2" style={{ borderTop: "1px solid var(--hf-border)" }}>
        <div>
          <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>
            {t.labelCrimpO}
          </div>
          {item.assemblyType === "stage" && item.crimp2 != null ? (
            <div className="text-lg font-extrabold leading-tight" style={{ color: "var(--hf-red)" }}>
              {stageCrimpDiameters(item).join(" → ")} mm
            </div>
          ) : (
            <div className="text-2xl font-extrabold leading-none" style={{ color: "var(--hf-red)" }}>
              {item.crimp.toFixed(1)} mm
            </div>
          )}
          <div className="hf-caption">{item.tol ? t.toleranceLabel(item.tol) : t.fixedSizeViaDie}</div>
        </div>
        {item.mandrel && (
          <div className="text-right">
            <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>
              {t.labelMandrel}
            </div>
            <div className="font-bold" style={{ color: "var(--hf-darkgrey)" }}>
              {item.mandrel}
            </div>
          </div>
        )}
      </div>

      <div className="flex items-center justify-between text-xs mt-1" style={{ color: "#999999" }}>
        <span>D1 {item.d1} · LF {item.lf}</span>
        <span className="hf-caption">{t.revShort(item.revDate, item.revSeq)}</span>
      </div>

      <button
        onClick={() => onOpenDetail(item)}
        className="hf-btn-outline mt-1 rounded px-3 py-2.5 sm:py-2 text-sm font-semibold flex items-center justify-center gap-1"
      >
        {t.btnDetails} <ChevronRight size={16} />
      </button>
    </div>
  );
}

function DetailModal({ item, isFav, onToggleFav, onClose, onPrint, onEdit, onOpenOnePiece, onOpenStageCrimp, onOpenStainlessInfo, t, lang }) {
  const steps = getSteps(item, t, lang);
  const note = itemNote(item, lang);
  return (
    <div
      className="screen-only fixed inset-0 z-50 flex items-end sm:items-center justify-center"
      style={{ background: "rgba(0,0,0,0.45)" }}
      onClick={onClose}
    >
      <div
        className="bg-white w-full sm:max-w-lg sm:rounded-lg max-h-[92vh] overflow-y-auto"
        style={{ border: "1px solid var(--hf-border)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 py-3" style={{ background: "var(--hf-red)" }}>
          <div>
            <div className="hf-headline text-white" style={{ fontSize: "18px" }}>
              {item.series}
            </div>
            <div className="text-white text-xs opacity-90">{item.standard} · {variantShortLabel(item.uitvoering, t)}</div>
          </div>
          <button onClick={onClose} aria-label={t.ariaClose} className="text-white flex-shrink-0 p-1">
            <X size={22} />
          </button>
        </div>

        <div className="p-4 flex flex-col gap-4">
          <div className="grid grid-cols-2 gap-3 p-3 rounded" style={{ background: "var(--hf-bg-light)" }}>
            <div>
              <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>{t.fieldNominalDiameter}</div>
              <div className="font-bold">DN {item.dn}</div>
            </div>
            <div>
              <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>{t.fieldFerruleMarking}</div>
              <div className="font-bold">{item.ferrule}</div>
            </div>
            <div>
              <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>{t.labelCrimpO}</div>
              <div className="font-extrabold" style={{ color: "var(--hf-red)" }}>
                {item.assemblyType === "stage" && item.crimp2 != null
                  ? `${stageCrimpDiameters(item).join(" → ")} mm`
                  : `${item.crimp.toFixed(1)} mm ${item.tol ? `(${item.tol})` : ""}`}
              </div>
            </div>
            <div>
              <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>{t.fieldMandrel}</div>
              <div className="font-bold">{item.mandrel || t.notApplicable}</div>
            </div>
            <div>
              <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>{t.fieldD1LF}</div>
              <div className="font-bold">{item.d1} / {item.lf} mm</div>
            </div>
            <div>
              <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>{t.fieldMfrCodes}</div>
              <div className="font-bold">{item.mfrCodes.join(" · ")}</div>
            </div>
            {item.da != null && (
              <div>
                <div className="text-[11px]" style={{ color: "var(--hf-grey-dark)" }}>{t.fieldDA}</div>
                <div className="font-bold">{item.da} mm</div>
              </div>
            )}
          </div>

          {item.assemblyType === "interlock" && (
            <button
              onClick={onOpenOnePiece}
              className="flex items-start gap-2 p-3 rounded text-sm text-left"
              style={{ background: "var(--hf-bg-light)", border: "1px solid var(--hf-border)" }}
            >
              <Link2 size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
              <span>
                {t.onePieceLinkText} <span style={{ color: "var(--hf-red)", fontWeight: 600 }}>{t.btnViewOnePiece} →</span>
              </span>
            </button>
          )}

          {item.assemblyType === "stage" && (
            <button
              onClick={onOpenStageCrimp}
              className="flex items-start gap-2 p-3 rounded text-sm text-left"
              style={{ background: "var(--hf-bg-light)", border: "1px solid var(--hf-border)" }}
            >
              <Link2 size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
              <span>
                {t.stageCrimpLinkText} <span style={{ color: "var(--hf-red)", fontWeight: 600 }}>{t.btnViewStageCrimp} →</span>
              </span>
            </button>
          )}

          {item.material === "rvs" && (
            <div className="flex flex-col gap-2">
              <div className="flex items-start gap-2 p-3 rounded text-sm" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
                <AlertTriangle size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
                <span>{t.stainlessAutoNoteText}</span>
              </div>
              <button
                onClick={onOpenStainlessInfo}
                className="flex items-start gap-2 p-3 rounded text-sm text-left"
                style={{ background: "var(--hf-bg-light)", border: "1px solid var(--hf-border)" }}
              >
                <Link2 size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
                <span>
                  {t.stainlessLinkText} <span style={{ color: "var(--hf-red)", fontWeight: 600 }}>{t.btnViewStainlessInfo} →</span>
                </span>
              </button>
            </div>
          )}

          {note && (
            <div className="flex items-start gap-2 p-3 rounded text-sm" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
              <Info size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
              <span>{note}</span>
            </div>
          )}

          <div>
            <div className="hf-subhead mb-2">{t.headingInstructions}</div>
            <ol className="flex flex-col gap-2">
              {steps.map((s, i) => (
                <li key={i} className="flex gap-2 text-sm">
                  <span
                    className="flex-shrink-0 flex items-center justify-center rounded-full font-bold text-white"
                    style={{ background: "var(--hf-darkgrey)", width: 20, height: 20, fontSize: 11 }}
                  >
                    {i + 1}
                  </span>
                  <span>{s}</span>
                </li>
              ))}
            </ol>
          </div>

          <div className="hf-caption">{t.referenceLine(item.revDate, item.revSeq)}</div>

          <div className="flex gap-2 pt-2 flex-wrap" style={{ borderTop: "1px solid var(--hf-border)" }}>
            <button
              onClick={() => onToggleFav(item.id)}
              className="hf-btn-outline flex-1 rounded px-3 py-2.5 text-sm font-semibold flex items-center justify-center gap-1"
            >
              <Star size={16} fill={isFav ? "#d7102d" : "none"} color={isFav ? "#d7102d" : "#34393b"} />
              {isFav ? t.btnRemoveFavorite : t.btnSaveFavorite}
            </button>
            <button
              onClick={() => onEdit(item)}
              className="hf-btn-outline flex-1 rounded px-3 py-2.5 text-sm font-semibold flex items-center justify-center gap-1"
            >
              <Pencil size={16} /> {t.btnEdit}
            </button>
            <button
              onClick={onPrint}
              className="hf-btn-primary flex-1 rounded px-3 py-2.5 text-sm font-semibold flex items-center justify-center gap-1"
            >
              <Printer size={16} /> {t.btnPrintCard}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function UpdatePanel({ onClose, onSave, onReset, hasCustomData, meta, saving, storageOk, t, locale, data }) {
  const [rawRows, setRawRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [parseResult, setParseResult] = useState(null);
  const [localError, setLocalError] = useState("");
  const [saveError, setSaveError] = useState("");
  const [confirmReset, setConfirmReset] = useState(false);

  async function handleFileChange(e) {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    setLocalError("");
    setSaveError("");
    setParseResult(null);
    if (!/\.xlsx$/i.test(file.name)) {
      setLocalError(t.errFileType);
      setFileName("");
      setRawRows(null);
      return;
    }
    setFileName(file.name);
    try {
      const rows = await readXlsxFile(file);
      setRawRows(rows);
    } catch (e2) {
      setLocalError(t.fileReadError);
      setRawRows(null);
    }
  }

  function handleAnalyse() {
    setLocalError("");
    setSaveError("");
    const result = validateRows(rawRows, t);
    setParseResult(result);
  }

  async function handleConfirm() {
    if (!parseResult || parseResult.errors.length > 0) return;
    setSaveError("");
    const res = await onSave(parseResult.rows);
    if (!res.ok) setSaveError(res.message || t.saveFailedDefault);
  }

  const isValid = parseResult && parseResult.errors.length === 0 && parseResult.rows.length > 0;

  return (
    <div
      className="screen-only fixed inset-0 z-50 flex items-end sm:items-center justify-center"
      style={{ background: "rgba(0,0,0,0.45)" }}
      onClick={onClose}
    >
      <div
        className="bg-white w-full sm:max-w-xl sm:rounded-lg max-h-[92vh] overflow-y-auto"
        style={{ border: "1px solid var(--hf-border)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 py-3" style={{ background: "var(--hf-red)" }}>
          <div className="hf-headline text-white" style={{ fontSize: "18px" }}>{t.updateTitle}</div>
          <button onClick={onClose} aria-label={t.ariaClose} className="text-white flex-shrink-0 p-1">
            <X size={22} />
          </button>
        </div>

        <div className="p-4 flex flex-col gap-4">
          {!storageOk && (
            <div className="flex items-start gap-2 p-3 rounded text-sm" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
              <AlertTriangle size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
              <span>{t.storageWarning}</span>
            </div>
          )}

          <div className="text-sm" style={{ color: "var(--hf-darkgrey)" }}>
            {storageOk ? t.introShared : t.introLocal}
          </div>

          <div className="flex flex-wrap gap-2">
            <label className="hf-btn-outline rounded px-3 py-2.5 sm:py-2 text-sm font-semibold flex items-center gap-2 cursor-pointer">
              <Upload size={16} /> {t.btnChooseFile}
              <input
                type="file"
                accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                onChange={handleFileChange}
                className="hidden"
              />
            </label>
            <button
              onClick={() => downloadBinaryFile("hansaflex-crimp-sjabloon.xlsx", buildTemplateXlsx(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
              className="hf-btn-outline rounded px-3 py-2.5 sm:py-2 text-sm font-semibold flex items-center gap-2"
            >
              <Download size={16} /> {t.btnDownloadTemplate}
            </button>
            <button
              onClick={() => downloadBinaryFile("hansaflex-crimp-data.xlsx", buildDataXlsx(data), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
              className="hf-btn-outline rounded px-3 py-2.5 sm:py-2 text-sm font-semibold flex items-center gap-2"
            >
              <Download size={16} /> {t.btnDownloadData}
            </button>
          </div>
          {fileName && (
            <div className="text-xs flex items-center gap-1" style={{ color: "var(--hf-grey-dark)" }}>
              <FileText size={14} /> {fileName}
            </div>
          )}

          {localError && <div className="text-sm" style={{ color: "var(--hf-red)" }}>{localError}</div>}

          <button
            onClick={handleAnalyse}
            disabled={!rawRows || rawRows.length === 0}
            className="hf-btn-primary rounded px-3 py-2.5 sm:py-2 text-sm font-semibold disabled:opacity-40"
          >
            {t.btnAnalyse}
          </button>

          {parseResult && parseResult.errors.length > 0 && (
            <div className="p-3 rounded text-sm flex flex-col gap-1" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
              <div className="font-bold flex items-center gap-1" style={{ color: "var(--hf-red)" }}>
                <AlertTriangle size={16} /> {t.errorsFoundTitle(parseResult.errors.length)}
              </div>
              <ul className="list-disc pl-5">
                {parseResult.errors.slice(0, 12).map((e, i) => (
                  <li key={i}>{e}</li>
                ))}
              </ul>
              {parseResult.errors.length > 12 && <div className="hf-caption">{t.moreErrors(parseResult.errors.length - 12)}</div>}
            </div>
          )}

          {isValid && (
            <div className="flex flex-col gap-2">
              <div className="p-3 rounded text-sm flex items-center gap-2" style={{ background: "var(--hf-bg-light)", border: "1px solid var(--hf-border)" }}>
                <CheckCircle2 size={16} style={{ color: "var(--hf-red)" }} />
                <span><strong>{parseResult.rows.length}</strong> {t.validatedCountSuffix}</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-xs" style={{ borderCollapse: "collapse" }}>
                  <thead>
                    <tr style={{ background: "var(--hf-red)", color: "#ffffff" }}>
                      <th className="text-left p-1.5">{t.tableHeaderSeries}</th>
                      <th className="text-left p-1.5">{t.tableHeaderDn}</th>
                      <th className="text-left p-1.5">{t.tableHeaderHose}</th>
                      <th className="text-left p-1.5">{t.labelCrimpO}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {parseResult.rows.slice(0, 5).map((r, i) => (
                      <tr key={r.id} style={{ background: i % 2 === 1 ? "var(--hf-bg-light)" : "#ffffff" }}>
                        <td className="p-1.5" style={{ borderBottom: "1px solid var(--hf-border)" }}>{r.series}</td>
                        <td className="p-1.5" style={{ borderBottom: "1px solid var(--hf-border)" }}>{r.dn}</td>
                        <td className="p-1.5" style={{ borderBottom: "1px solid var(--hf-border)" }}>{r.ferrule}</td>
                        <td className="p-1.5" style={{ borderBottom: "1px solid var(--hf-border)" }}>{r.crimp} mm</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {parseResult.rows.length > 5 && <div className="hf-caption mt-1">{t.moreRows(parseResult.rows.length - 5)}</div>}
              </div>

              {saveError && <div className="text-sm" style={{ color: "var(--hf-red)" }}>{saveError}</div>}

              <div className="flex gap-2">
                <button
                  onClick={() => setParseResult(null)}
                  className="hf-btn-outline flex-1 rounded px-3 py-2.5 text-sm font-semibold"
                >
                  {t.btnCancel}
                </button>
                <button
                  onClick={handleConfirm}
                  disabled={saving}
                  className="hf-btn-primary flex-1 rounded px-3 py-2.5 text-sm font-semibold disabled:opacity-60"
                >
                  {saving ? t.savingLabel : t.btnConfirmSave}
                </button>
              </div>
            </div>
          )}

          <div className="pt-3 flex flex-col gap-2" style={{ borderTop: "1px solid var(--hf-border)" }}>
            <div className="hf-caption">
              {meta ? t.statusUpdatedOn(new Date(meta.updatedAt).toLocaleString(locale), meta.rowCount) : t.statusNoUpload}
            </div>
            {hasCustomData && (
              <button
                onClick={() => {
                  if (!confirmReset) { setConfirmReset(true); return; }
                  onReset();
                  setConfirmReset(false);
                }}
                className="hf-btn-outline self-start rounded px-3 py-2 text-xs font-semibold flex items-center gap-1"
                style={confirmReset ? { borderColor: "var(--hf-red)", color: "var(--hf-red)" } : {}}
              >
                <RotateCcw size={14} /> {confirmReset ? t.btnResetConfirm : t.btnResetDefault}
              </button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Test-Doorn (mandrel) - naslagpagina met eigen schematische SVG's   */
/* ------------------------------------------------------------------ */
function MandrelFigureSvg({ variant }) {
  const width = 260;
  const height = 70;
  const bodyX = 96;
  const bodyW = width - bodyX - 10;
  const depthFraction = variant === 1 ? 0.35 : variant === 2 ? 0.55 : variant === 3 ? 0.5 : 1.0;
  const insertW = bodyW * depthFraction;
  const overshoot = variant === 4;
  return (
    <svg viewBox={`0 0 ${width} ${height}`} width="100%" height={height} style={{ maxWidth: 280, display: "block" }}>
      <rect x={bodyX} y={8} width={bodyW} height={54} rx={4} fill="#ecf1f4" stroke="#34393b" strokeWidth="1.5" />
      {variant === 3 && (
        <rect x={bodyX} y={8} width={bodyW} height={54} rx={4} fill="none" stroke="#d7102d" strokeWidth="1.5" strokeDasharray="3 2" />
      )}
      <rect x={10} y={26} width={bodyX - 10 + insertW} height={18} rx={3} fill="#ffffff" stroke="#34393b" strokeWidth="1.5" />
      <line x1={bodyX - 10 + insertW} y1={6} x2={bodyX - 10 + insertW} y2={64} stroke="#d7102d" strokeWidth="2" strokeDasharray={overshoot ? "0" : "4 3"} />
      {overshoot && (
        <>
          <line x1={width - 14} y1={14} x2={width - 4} y2={58} stroke="#d7102d" strokeWidth="2" />
          <line x1={width - 4} y1={14} x2={width - 14} y2={58} stroke="#d7102d" strokeWidth="2" />
        </>
      )}
      <text x={10} y={20} fontSize="9" fill="#666666">HANSA-FLEX</text>
    </svg>
  );
}

function TestMandrelPage({ t }) {
  return (
    <div className="flex flex-col gap-4 max-w-3xl">
      <div className="flex items-start gap-2 p-3 rounded text-sm" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
        <AlertTriangle size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
        <span className="font-semibold">{t.mandrelPageWarning}</span>
      </div>
      <p className="text-sm">{t.mandrelPageIntro}</p>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        {t.mandrelFigures.map((fig, i) => (
          <div key={i} className="hf-card p-3 flex flex-col gap-2">
            <div className="flex items-center gap-1.5">
              {i < 2 ? (
                <CheckCircle2 size={15} style={{ color: "var(--hf-red)" }} />
              ) : (
                <AlertTriangle size={15} style={{ color: "var(--hf-red)" }} />
              )}
              <span className="hf-subhead" style={{ fontSize: 13 }}>{fig.title}</span>
            </div>
            <MandrelFigureSvg variant={i + 1} />
            <p className="text-xs" style={{ color: "var(--hf-darkgrey)" }}>{fig.text}</p>
          </div>
        ))}
      </div>

      <div>
        <div className="hf-subhead mb-1">{t.mandrelFrequencyHeading}</div>
        <p className="text-sm">{t.mandrelFrequencyText}</p>
      </div>
      <div>
        <div className="hf-subhead mb-1">{t.mandrelControlHeading}</div>
        <p className="text-sm">{t.mandrelControlText}</p>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  One Piece Fitting - naslagpagina met eigen schematische SVG        */
/* ------------------------------------------------------------------ */
function OnePieceFittingSvg() {
  return (
    <svg viewBox="0 0 320 80" width="100%" height={90} style={{ maxWidth: 340, display: "block" }}>
      <rect x="10" y="34" width="150" height="14" rx="3" fill="#ffffff" stroke="#34393b" strokeWidth="1.5" />
      <rect x="150" y="20" width="60" height="42" rx="4" fill="#ecf1f4" stroke="#34393b" strokeWidth="1.5" />
      <rect x="205" y="28" width="45" height="26" rx="3" fill="#ffffff" stroke="#34393b" strokeWidth="1.5" />
      <line x1="150" y1="14" x2="150" y2="68" stroke="#d7102d" strokeWidth="2" strokeDasharray="4 3" />
      <text x="118" y="10" fontSize="9" fill="#666666" textAnchor="middle">insteekdiepte</text>
      <text x="30" y="30" fontSize="9" fill="#666666">HANSA-FLEX</text>
    </svg>
  );
}

function OnePieceFittingPage({ t }) {
  return (
    <div className="flex flex-col gap-4 max-w-3xl">
      <p className="text-sm">{t.onePiecePageIntro}</p>
      <div className="hf-card p-3">
        <OnePieceFittingSvg />
      </div>
      <ol className="flex flex-col gap-2">
        {t.onePieceSteps.map((s, i) => (
          <li key={i} className="flex gap-2 text-sm">
            <span
              className="flex-shrink-0 flex items-center justify-center rounded-full font-bold text-white"
              style={{ background: "var(--hf-darkgrey)", width: 20, height: 20, fontSize: 11 }}
            >
              {i + 1}
            </span>
            <span>{s}</span>
          </li>
        ))}
      </ol>
      <div className="flex items-start gap-2 p-3 rounded text-sm" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
        <AlertTriangle size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
        <span>{t.onePieceWarningText}</span>
      </div>
      <p className="hf-caption">{t.onePieceNoteText}</p>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Interlock Fitting - naslagpagina met eigen schematische SVG        */
/*  (perslijst p.30: "Assembly instruction - interlock fittings")      */
/* ------------------------------------------------------------------ */
function InterlockFittingSvg() {
  return (
    <svg viewBox="0 0 320 90" width="100%" height={95} style={{ maxWidth: 340, display: "block" }}>
      <rect x="10" y="38" width="140" height="14" rx="3" fill="#ffffff" stroke="#34393b" strokeWidth="1.5" />
      <rect x="145" y="24" width="55" height="42" rx="4" fill="#ecf1f4" stroke="#34393b" strokeWidth="1.5" />
      <rect x="196" y="30" width="20" height="30" rx="2" fill="#ffffff" stroke="#d7102d" strokeWidth="1.5" />
      <rect x="212" y="24" width="55" height="42" rx="4" fill="#ecf1f4" stroke="#34393b" strokeWidth="1.5" />
      <line x1="196" y1="16" x2="196" y2="76" stroke="#d7102d" strokeWidth="2" strokeDasharray="4 3" />
      <line x1="216" y1="16" x2="216" y2="76" stroke="#d7102d" strokeWidth="2" strokeDasharray="4 3" />
      <text x="206" y="12" fontSize="8" fill="#666666" textAnchor="middle">markering</text>
      <text x="30" y="34" fontSize="9" fill="#666666">HANSA-FLEX</text>
    </svg>
  );
}

function InterlockFittingPage({ t }) {
  return (
    <div className="flex flex-col gap-4 max-w-3xl">
      <p className="text-sm">{t.interlockPageIntro}</p>
      <div className="hf-card p-3">
        <InterlockFittingSvg />
      </div>
      <ol className="flex flex-col gap-2">
        {t.interlockSteps.map((s, i) => (
          <li key={i} className="flex gap-2 text-sm">
            <span
              className="flex-shrink-0 flex items-center justify-center rounded-full font-bold text-white"
              style={{ background: "var(--hf-darkgrey)", width: 20, height: 20, fontSize: 11 }}
            >
              {i + 1}
            </span>
            <span>{s}</span>
          </li>
        ))}
      </ol>
      <div className="flex items-start gap-2 p-3 rounded text-sm" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
        <AlertTriangle size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
        <span>{t.interlockWarningText}</span>
      </div>
      <p className="hf-caption">{t.interlockNoteText}</p>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Trapsgewijs persen - naslagpagina met eigen schematische SVG       */
/*  (perslijst staal p.45 "SGB 100 - stage crimping", RVS p.21)        */
/* ------------------------------------------------------------------ */
function StageCrimpSvg() {
  return (
    <svg viewBox="0 0 320 90" width="100%" height={95} style={{ maxWidth: 340, display: "block" }}>
      <rect x="10" y="38" width="120" height="14" rx="3" fill="#ffffff" stroke="#34393b" strokeWidth="1.5" />
      <rect x="130" y="10" width="70" height="70" rx="6" fill="none" stroke="#d1dce3" strokeWidth="2" strokeDasharray="3 2" />
      <rect x="140" y="18" width="60" height="54" rx="5" fill="none" stroke="#999999" strokeWidth="2" strokeDasharray="3 2" />
      <rect x="150" y="26" width="50" height="38" rx="4" fill="#ecf1f4" stroke="#d7102d" strokeWidth="2" />
      <text x="230" y="20" fontSize="9" fill="#666666">Ø1</text>
      <text x="230" y="45" fontSize="9" fill="#666666">Ø2</text>
      <text x="230" y="68" fontSize="9" fill="#d7102d" fontWeight="700">Ø3 (eindmaat)</text>
      <text x="30" y="34" fontSize="9" fill="#666666">HANSA-FLEX</text>
    </svg>
  );
}

function StageCrimpPage({ t, exampleItem }) {
  const diameters = exampleItem ? stageCrimpDiameters(exampleItem) : null;
  return (
    <div className="flex flex-col gap-4 max-w-3xl">
      <p className="text-sm">{t.stageCrimpPageIntro}</p>
      <div className="hf-card p-3">
        <StageCrimpSvg />
      </div>
      <div>
        <div className="hf-subhead mb-1">{t.stageCrimpWhyHeading}</div>
        <p className="text-sm">{t.stageCrimpWhyText}</p>
      </div>
      <div>
        <div className="hf-subhead mb-1">{t.stageCrimpExampleHeading}</div>
        <p className="text-sm">{t.stageCrimpExampleText}</p>
        {diameters && (
          <div className="flex items-center gap-2 mt-2 flex-wrap">
            {diameters.map((d, i) => (
              <React.Fragment key={i}>
                {i > 0 && <ChevronRight size={16} style={{ color: "#999999" }} />}
                <span
                  className="rounded px-2.5 py-1 text-sm font-bold"
                  style={{
                    background: i === diameters.length - 1 ? "var(--hf-red)" : "var(--hf-bg-light)",
                    color: i === diameters.length - 1 ? "#ffffff" : "var(--hf-darkgrey)",
                    border: "1px solid var(--hf-border)",
                  }}
                >
                  Ø{i + 1} {d} mm
                </span>
              </React.Fragment>
            ))}
          </div>
        )}
      </div>
      <div>
        <div className="hf-subhead mb-1">{t.stageCrimpApplicableHeading}</div>
        <p className="text-sm">{t.stageCrimpApplicableText}</p>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  RVS-fittingen - naslagpagina met materiaalregels (crimp chart VA)  */
/* ------------------------------------------------------------------ */
function StainlessPage({ t }) {
  return (
    <div className="flex flex-col gap-4 max-w-3xl">
      <p className="text-sm">{t.stainlessPageIntro}</p>

      <div className="flex items-start gap-2 p-3 rounded text-sm" style={{ background: "#fdeceb", border: "1px solid var(--hf-red)" }}>
        <AlertTriangle size={16} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
        <div>
          <div className="font-semibold">{t.stainlessHoldTimeHeading}</div>
          <div>{t.stainlessHoldTimeText}</div>
        </div>
      </div>

      <div>
        <div className="hf-subhead mb-1">{t.stainlessDisclaimerHeading}</div>
        <p className="text-sm">{t.stainlessDisclaimerText}</p>
      </div>

      <div>
        <div className="hf-subhead mb-1">{t.stainlessInterlockDiffHeading}</div>
        <p className="text-sm">{t.stainlessInterlockDiffText}</p>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Wijzigen: combinatie toevoegen / bewerken / verwijderen            */
/* ------------------------------------------------------------------ */
function MfrCodeEditor({ codes, onChange, placeholder }) {
  const [draft, setDraft] = useState("");
  function commit() {
    const v = draft.trim();
    if (v && !codes.includes(v)) onChange([...codes, v]);
    setDraft("");
  }
  return (
    <div>
      <div className="flex flex-wrap gap-1.5 mb-1.5">
        {codes.map((c) => (
          <span key={c} className="hf-chip px-2 py-1 rounded flex items-center gap-1">
            {c}
            <button type="button" onClick={() => onChange(codes.filter((x) => x !== c))} aria-label="remove">
              <X size={12} />
            </button>
          </span>
        ))}
      </div>
      <input
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" || e.key === ",") {
            e.preventDefault();
            commit();
          }
        }}
        onBlur={commit}
        placeholder={placeholder}
        className="w-full rounded p-2 text-sm"
        style={{ border: "1px solid var(--hf-border)" }}
      />
    </div>
  );
}

function EditModal({ item, isNew, onClose, onSave, onDelete, saving, t }) {
  const [form, setForm] = useState({
    ...item,
    dn: item.dn ?? "",
    crimp: item.crimp ?? "",
    crimp2: item.crimp2 ?? "",
    crimp3: item.crimp3 ?? "",
    d1: item.d1 ?? "",
    lf: item.lf ?? "",
    da: item.da ?? "",
    tol: item.tol ?? "",
    mandrel: item.mandrel ?? "",
    skiveLength: item.skiveLength ?? "",
    skiveLengthInt: item.skiveLengthInt ?? "",
    note: item.note ?? "",
    assemblyType: item.assemblyType || "standard",
    material: item.material || "",
  });
  const [error, setError] = useState("");
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [savingLocal, setSavingLocal] = useState(false);

  function set(field, value) {
    setForm((f) => ({ ...f, [field]: value }));
  }

  function buildPayload() {
    return {
      id: form.id,
      series: form.series.trim(),
      standard: form.standard.trim(),
      uitvoering: form.uitvoering,
      uitvoeringLabel: form.uitvoering === "geskived" ? "geskived" : "geen skive",
      dn: Number(form.dn),
      mfrCodes: form.mfrCodes,
      ferrule: form.ferrule.trim(),
      crimp: Number(form.crimp),
      crimp2: form.crimp2 !== "" ? Number(form.crimp2) : undefined,
      crimp3: form.crimp3 !== "" ? Number(form.crimp3) : undefined,
      tol: form.tol ? form.tol.trim() : null,
      mandrel: form.mandrel ? form.mandrel.trim() : null,
      d1: Number(form.d1),
      da: form.da !== "" ? Number(form.da) : undefined,
      lf: Number(form.lf),
      skiveLength: form.skiveLength !== "" ? Number(form.skiveLength) : undefined,
      skiveLengthInt: form.skiveLengthInt !== "" ? Number(form.skiveLengthInt) : undefined,
      assemblyType: form.assemblyType,
      material: form.material === "rvs" ? "rvs" : undefined,
      note: form.note ? form.note.trim() : undefined,
    };
  }

  function isValid() {
    return (
      form.series.trim() &&
      form.standard.trim() &&
      form.dn !== "" && !Number.isNaN(Number(form.dn)) &&
      form.mfrCodes.length > 0 &&
      form.ferrule.trim() &&
      form.crimp !== "" && !Number.isNaN(Number(form.crimp)) &&
      form.d1 !== "" && !Number.isNaN(Number(form.d1)) &&
      form.lf !== "" && !Number.isNaN(Number(form.lf))
    );
  }

  async function handleSave() {
    if (!isValid()) {
      setError(t.editErrorRequired);
      return;
    }
    setError("");
    setSavingLocal(true);
    const res = await onSave(buildPayload());
    setSavingLocal(false);
    if (!res.ok) setError(res.message || t.saveFailedDefault);
  }

  async function handleDeleteClick() {
    if (!confirmDelete) {
      setConfirmDelete(true);
      return;
    }
    setSavingLocal(true);
    await onDelete(form.id);
    setSavingLocal(false);
  }

  return (
    <div
      className="screen-only fixed inset-0 z-50 flex items-end sm:items-center justify-center"
      style={{ background: "rgba(0,0,0,0.45)" }}
      onClick={onClose}
    >
      <div
        className="bg-white w-full sm:max-w-xl sm:rounded-lg max-h-[92vh] overflow-y-auto"
        style={{ border: "1px solid var(--hf-border)" }}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="flex items-center justify-between px-4 py-3" style={{ background: "var(--hf-red)" }}>
          <div className="hf-headline text-white" style={{ fontSize: "18px" }}>
            {isNew ? t.editTitleNew : t.editTitleEdit}
          </div>
          <button onClick={onClose} aria-label={t.ariaClose} className="text-white flex-shrink-0 p-1">
            <X size={22} />
          </button>
        </div>

        <div className="p-4 flex flex-col gap-3">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldSeries}</label>
              <input value={form.series} onChange={(e) => set("series", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldStandard}</label>
              <input value={form.standard} onChange={(e) => set("standard", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldUitvoering}</label>
              <select value={form.uitvoering} onChange={(e) => set("uitvoering", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }}>
                <option value="standaard">{t.variantStandardChip}</option>
                <option value="geskived">{t.variantSkivedChip}</option>
              </select>
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldDn}</label>
              <input type="number" value={form.dn} onChange={(e) => set("dn", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
          </div>

          <div>
            <label className="hf-subhead block mb-1">{t.editFieldMfrCodes}</label>
            <MfrCodeEditor codes={form.mfrCodes} onChange={(codes) => set("mfrCodes", codes)} placeholder={t.editFieldMfrAdd} />
          </div>

          <div>
            <label className="hf-subhead block mb-1">{t.editFieldFerrule}</label>
            <input value={form.ferrule} onChange={(e) => set("ferrule", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldCrimp}</label>
              <input type="number" step="0.1" value={form.crimp} onChange={(e) => set("crimp", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldCrimp2}</label>
              <input type="number" step="0.1" value={form.crimp2} onChange={(e) => set("crimp2", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldCrimp3}</label>
              <input type="number" step="0.1" value={form.crimp3} onChange={(e) => set("crimp3", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldTol}</label>
              <input value={form.tol} onChange={(e) => set("tol", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldMandrel}</label>
              <input value={form.mandrel} onChange={(e) => set("mandrel", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldD1}</label>
              <input type="number" step="0.1" value={form.d1} onChange={(e) => set("d1", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldDa}</label>
              <input type="number" step="0.1" value={form.da} onChange={(e) => set("da", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldLf}</label>
              <input type="number" step="0.1" value={form.lf} onChange={(e) => set("lf", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldSkive}</label>
              <input type="number" step="0.1" value={form.skiveLength} onChange={(e) => set("skiveLength", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldSkiveInt}</label>
              <input type="number" step="0.1" value={form.skiveLengthInt} onChange={(e) => set("skiveLengthInt", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldAssemblyType}</label>
              <select value={form.assemblyType} onChange={(e) => set("assemblyType", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }}>
                <option value="standard">{t.assemblyTypeStandard}</option>
                <option value="interlock">{t.assemblyTypeInterlock}</option>
                <option value="stage">{t.assemblyTypeStage}</option>
              </select>
            </div>
            <div>
              <label className="hf-subhead block mb-1">{t.editFieldMaterial}</label>
              <select value={form.material} onChange={(e) => set("material", e.target.value)} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }}>
                <option value="">{t.materialLabelStaal}</option>
                <option value="rvs">{t.materialLabelRvs}</option>
              </select>
            </div>
          </div>

          <div>
            <label className="hf-subhead block mb-1">{t.editFieldNote}</label>
            <textarea value={form.note} onChange={(e) => set("note", e.target.value)} rows={2} className="w-full rounded p-2 text-sm" style={{ border: "1px solid var(--hf-border)" }} />
          </div>

          {error && <div className="text-sm" style={{ color: "var(--hf-red)" }}>{error}</div>}

          <div className="flex gap-2 pt-2 flex-wrap" style={{ borderTop: "1px solid var(--hf-border)" }}>
            <button onClick={onClose} className="hf-btn-outline flex-1 rounded px-3 py-2.5 text-sm font-semibold">
              {t.btnCancelEdit}
            </button>
            <button
              onClick={handleSave}
              disabled={savingLocal || saving}
              className="hf-btn-primary flex-1 rounded px-3 py-2.5 text-sm font-semibold flex items-center justify-center gap-1 disabled:opacity-60"
            >
              <Save size={16} /> {t.btnSaveEdit}
            </button>
          </div>
          {!isNew && (
            <button
              onClick={handleDeleteClick}
              disabled={savingLocal || saving}
              className="rounded px-3 py-2 text-xs font-semibold self-start flex items-center gap-1"
              style={{ color: "var(--hf-red)", border: "1px solid var(--hf-red)", background: "#ffffff" }}
            >
              <Trash2 size={14} /> {confirmDelete ? t.btnDeleteConfirm : t.btnDeleteItem}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Hoofdcomponent                                                    */
/* ------------------------------------------------------------------ */
export default function CrimpApp() {
  const [lang, setLang] = useState("nl");
  const t = I18N[lang];
  const locale = LOCALE_MAP[lang];

  const [seriesQuery, setSeriesQuery] = useState("");
  const [selectedSeries, setSelectedSeries] = useState(null);
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [dnFilter, setDnFilter] = useState("");
  const [mfrFilter, setMfrFilter] = useState("");
  const [variantFilter, setVariantFilter] = useState("alle");
  const [favorites, setFavorites] = useState([]);
  const [selectedItem, setSelectedItem] = useState(null);
  const [activeTab, setActiveTab] = useState("zoeken");

  const [customData, setCustomData] = useState(null);
  const [dataMeta, setDataMeta] = useState(null);
  const [storageOk, setStorageOk] = useState(true);
  const [showUpdatePanel, setShowUpdatePanel] = useState(false);
  const [saving, setSaving] = useState(false);

  const [editingItem, setEditingItem] = useState(null);
  const [isNewItem, setIsNewItem] = useState(false);

  // Bij het openen: kijk of er via "Data bijwerken" al een gedeelde dataset is
  // opgeslagen. Zo niet (of geen opslag beschikbaar), gebruik de ingebouwde
  // voorbeelddata — de tool blijft dus altijd werken.
  useEffect(() => {
    let cancelled = false;
    async function load() {
      if (!window.storage) {
        if (!cancelled) setStorageOk(false);
        return;
      }
      try {
        const res = await window.storage.get("crimp-data", true);
        if (!cancelled && res && res.value) {
          const parsed = JSON.parse(res.value);
          if (Array.isArray(parsed) && parsed.length > 0) setCustomData(parsed);
        }
      } catch (e) {
        // nog geen upload aanwezig — dat is geen fout, gewoon voorbeelddata gebruiken
      }
      try {
        const metaRes = await window.storage.get("crimp-data-meta", true);
        if (!cancelled && metaRes && metaRes.value) setDataMeta(JSON.parse(metaRes.value));
      } catch (e) {
        // idem
      }
    }
    load();
    return () => { cancelled = true; };
  }, []);

  const activeData = customData !== null ? customData : DEFAULT_DATA;

  const seriesList = useMemo(() => Array.from(new Set(activeData.map((d) => d.series))), [activeData]);

  const seriesSuggestions = useMemo(() => {
    if (!seriesQuery) return [];
    const q = seriesQuery.toLowerCase();
    return seriesList.filter((s) => s.toLowerCase().includes(q)).slice(0, 6);
  }, [seriesQuery, seriesList]);

  const availableDns = useMemo(() => {
    const pool = selectedSeries ? activeData.filter((d) => d.series === selectedSeries) : activeData;
    return Array.from(new Set(pool.map((d) => d.dn))).sort((a, b) => a - b);
  }, [selectedSeries, activeData]);

  const availableMfrCodes = useMemo(() => {
    let pool = activeData;
    if (selectedSeries) pool = pool.filter((d) => d.series === selectedSeries);
    if (dnFilter) pool = pool.filter((d) => String(d.dn) === String(dnFilter));
    const codes = new Set();
    pool.forEach((d) => d.mfrCodes.forEach((c) => codes.add(c)));
    return Array.from(codes).sort();
  }, [selectedSeries, dnFilter, activeData]);

  const filteredResults = useMemo(() => {
    return activeData.filter((d) => {
      if (selectedSeries && d.series !== selectedSeries) return false;
      if (!selectedSeries && seriesQuery) {
        const q = seriesQuery.toLowerCase();
        if (!d.series.toLowerCase().includes(q) && !d.ferrule.toLowerCase().includes(q)) return false;
      }
      if (dnFilter && String(d.dn) !== String(dnFilter)) return false;
      if (mfrFilter && !d.mfrCodes.includes(mfrFilter)) return false;
      if (variantFilter !== "alle" && d.uitvoering !== variantFilter) return false;
      return true;
    });
  }, [selectedSeries, seriesQuery, dnFilter, mfrFilter, variantFilter, activeData]);

  const favoriteItems = activeData.filter((d) => favorites.includes(d.id));

  function toggleFavorite(id) {
    setFavorites((f) => (f.includes(id) ? f.filter((x) => x !== id) : [...f, id]));
  }

  function resetFilters() {
    setSeriesQuery("");
    setSelectedSeries(null);
    setDnFilter("");
    setMfrFilter("");
    setVariantFilter("alle");
  }

  function selectSeries(s) {
    setSelectedSeries(s);
    setSeriesQuery(s);
    setDnFilter("");
    setMfrFilter("");
    setShowSuggestions(false);
  }

  function handlePrint() {
    window.print();
  }

  // Slaat een volledige, nieuwe rijenset op (gedeeld indien mogelijk, anders
  // alleen lokaal voor dit scherm) — gebruikt door zowel de xlsx-upload als
  // door de handmatige Wijzigen/toevoegen/verwijderen-acties.
  async function persistData(rows) {
    setSaving(true);
    try {
      if (window.storage) {
        const metaObj = { updatedAt: new Date().toISOString(), rowCount: rows.length };
        await window.storage.set("crimp-data", JSON.stringify(rows), true);
        await window.storage.set("crimp-data-meta", JSON.stringify(metaObj), true);
        setDataMeta(metaObj);
      }
      setCustomData(rows);
      setSaving(false);
      return { ok: true };
    } catch (e) {
      setSaving(false);
      return { ok: false, message: e.message || t.saveFailedDefault };
    }
  }

  // Bulk-upload (.xlsx): per rij wordt eerst gecontroleerd of de inhoud daadwerkelijk
  // afwijkt van de huidige combinatie met hetzelfde id. Is er niets veranderd, dan
  // blijft de bestaande revisie staan. Is er wel een wijziging (of is het een nieuwe
  // combinatie), dan wordt de revisie met 1 opgehoogd (of gestart op R01).
  async function handleSaveUploadedData(rows) {
    const byId = new Map(activeData.map((d) => [d.id, d]));
    const revisedRows = rows.map((r) => {
      const existing = byId.get(r.id);
      if (existing && rowsContentEqual(existing, r)) {
        return { ...r, revDate: existing.revDate, revSeq: existing.revSeq || 1 };
      }
      return { ...r, ...bumpRevision(existing) };
    });
    return persistData(revisedRows);
  }

  async function handleResetToDefault() {
    setSaving(true);
    if (window.storage) {
      try {
        await window.storage.delete("crimp-data", true);
      } catch (e) {}
      try {
        await window.storage.delete("crimp-data-meta", true);
      } catch (e) {}
    }
    setCustomData(null);
    setDataMeta(null);
    setSaving(false);
    resetFilters();
  }

  function openEditItem(item) {
    setEditingItem(item);
    setIsNewItem(false);
  }

  function openAddItem() {
    setEditingItem({
      id: "",
      series: "",
      standard: "",
      uitvoering: "standaard",
      uitvoeringLabel: "geen skive",
      dn: "",
      mfrCodes: [],
      ferrule: "",
      crimp: "",
      tol: null,
      mandrel: null,
      d1: "",
      da: null,
      lf: "",
      skiveLength: null,
      skiveLengthInt: null,
      assemblyType: "standard",
      note: "",
    });
    setIsNewItem(true);
  }

  function closeEdit() {
    setEditingItem(null);
    setIsNewItem(false);
  }

  async function handleSaveEdit(formData) {
    const existingIds = new Set(activeData.map((d) => d.id));
    let id = formData.id;
    if (isNewItem || !id) {
      const base = `${slugify(formData.series) || "combi"}-dn${formData.dn}-${slugify(formData.ferrule) || "x"}`;
      id = base;
      let n = 2;
      while (existingIds.has(id)) {
        id = `${base}-${n}`;
        n += 1;
      }
    }
    const existing = activeData.find((d) => d.id === formData.id);
    const revised = { ...formData, id, ...bumpRevision(existing) };
    const newRows = isNewItem || !existing
      ? [...activeData, revised]
      : activeData.map((d) => (d.id === formData.id ? revised : d));
    const res = await persistData(newRows);
    if (res.ok) {
      closeEdit();
      if (selectedItem && selectedItem.id === formData.id) setSelectedItem(revised);
    }
    return res;
  }

  async function handleDeleteItem(id) {
    const newRows = activeData.filter((d) => d.id !== id);
    setFavorites((f) => f.filter((x) => x !== id));
    const res = await persistData(newRows);
    if (res.ok) {
      closeEdit();
      if (selectedItem && selectedItem.id === id) setSelectedItem(null);
    }
    return res;
  }

  const printItems = selectedItem ? [selectedItem] : activeTab === "favorieten" ? favoriteItems : filteredResults;
  const printLabel = selectedItem ? t.printLabelOne : activeTab === "favorieten" ? t.printLabelFavorites : t.printLabelResults;

  return (
    <div className="hfapp min-h-screen" style={{ background: "#ffffff" }} lang={lang}>
      <style>{`
        .hfapp {
          --hf-red: #d7102d;
          --hf-white: #ffffff;
          --hf-black: #000000;
          --hf-darkgrey: #34393b;
          --hf-bg-light: #ecf1f4;
          --hf-border: #d1dce3;
          --hf-grey: #999999;
          --hf-grey-dark: #666666;
          font-family: Roboto, Calibri, system-ui, -apple-system, sans-serif;
          color: #000000;
        }
        .hf-headline {
          font-family: "Segoe UI", Gilroy, system-ui, -apple-system, sans-serif;
          font-weight: 800;
          text-transform: uppercase;
          letter-spacing: 0.02em;
        }
        .hf-subhead {
          font-family: Roboto, system-ui, sans-serif;
          font-weight: 700;
          font-size: 14px;
          color: var(--hf-darkgrey);
        }
        .hf-caption {
          font-style: italic;
          font-size: 12px;
          color: var(--hf-grey);
        }
        .hf-card {
          background: #ffffff;
          border: 1px solid var(--hf-border);
          border-radius: 8px;
          transition: box-shadow 0.15s ease, border-color 0.15s ease;
        }
        .hf-card:hover {
          border-color: var(--hf-red);
          box-shadow: 0 4px 14px rgba(0, 0, 0, 0.08);
        }
        .hf-btn-primary {
          background: var(--hf-red);
          color: #ffffff;
          border: 1px solid var(--hf-red);
        }
        .hf-btn-primary:hover {
          background: #b60d26;
        }
        .hf-btn-outline {
          background: #ffffff;
          color: var(--hf-darkgrey);
          border: 1px solid var(--hf-border);
        }
        .hf-btn-outline:hover {
          border-color: var(--hf-red);
          color: var(--hf-red);
        }
        .hf-chip {
          border: 1px solid var(--hf-border);
          background: var(--hf-bg-light);
          color: var(--hf-darkgrey);
          font-size: 12px;
          font-weight: 600;
        }
        .hf-chip.active {
          background: var(--hf-red);
          border-color: var(--hf-red);
          color: #ffffff;
        }
        .hf-tab {
          font-weight: 700;
          font-size: 14px;
          color: var(--hf-grey-dark);
          padding: 8px 4px;
          border-bottom: 3px solid transparent;
        }
        .hf-tab.active {
          color: var(--hf-red);
          border-bottom-color: var(--hf-red);
        }
        .print-only { display: none; }
        @media print {
          .screen-only { display: none !important; }
          .print-only { display: block !important; }
          body { background: #ffffff; }
        }
      `}</style>

      {/* Disclaimerbanner */}
      <div className="screen-only flex items-start gap-2 px-4 py-2" style={{ background: "#fdeceb", borderBottom: "2px solid var(--hf-red)" }}>
        <AlertTriangle size={18} style={{ color: "var(--hf-red)", flexShrink: 0, marginTop: 2 }} />
        <p className="text-xs sm:text-sm">
          <strong>{t.disclaimerStrong}</strong> {t.disclaimerRest}
        </p>
      </div>

      {/* Header */}
      <header className="screen-only flex items-center justify-between gap-3 px-4 py-3 border-b flex-wrap" style={{ borderColor: "var(--hf-border)" }}>
        <div className="flex items-center gap-3 sm:gap-4">
          <Logo height={30} />
          <div className="hidden sm:block" style={{ width: 1, alignSelf: "stretch", background: "var(--hf-border)" }} />
          <div className="flex flex-col leading-none">
            <span className="hf-headline" style={{ fontSize: "24px" }}>{t.appTitle}</span>
            <span className="hf-subhead" style={{ fontSize: "12px", marginTop: "3px", fontWeight: 600 }}>
              {t.appSubtitle}
            </span>
          </div>
        </div>
        <div className="flex items-center gap-2">
          <LanguageSwitcher lang={lang} onChange={setLang} label={t.languageLabel} />
          <button onClick={() => setShowUpdatePanel(true)} className="hf-btn-outline rounded px-3 py-2.5 sm:py-2 text-sm font-semibold flex items-center gap-2">
            <Upload size={16} /> {t.btnUpdateData}
          </button>
          <button onClick={handlePrint} className="hf-btn-primary rounded px-3 py-2.5 sm:py-2 text-sm font-semibold flex items-center gap-2">
            <Printer size={16} /> {t.btnPrintExport}
          </button>
        </div>
      </header>

      {/* Databronstatus */}
      <div className="screen-only flex items-center gap-1.5 px-4 py-1.5" style={{ background: "var(--hf-bg-light)", borderBottom: "1px solid var(--hf-border)" }}>
        <span
          style={{
            display: "inline-block",
            width: 7,
            height: 7,
            borderRadius: "50%",
            background: customData ? "var(--hf-red)" : "#999999",
          }}
        />
        <span className="hf-caption">
          {customData && dataMeta
            ? t.statusShared(new Date(dataMeta.updatedAt).toLocaleString(locale), dataMeta.rowCount)
            : t.statusDefault}
        </span>
      </div>

      {/* Tabs */}
      <nav className="screen-only flex gap-4 px-4 border-b overflow-x-auto" style={{ borderColor: "var(--hf-border)" }}>
        <button className={`hf-tab ${activeTab === "zoeken" ? "active" : ""}`} onClick={() => { setActiveTab("zoeken"); setSelectedItem(null); }}>
          {t.tabSearch}
        </button>
        <button className={`hf-tab ${activeTab === "favorieten" ? "active" : ""}`} onClick={() => { setActiveTab("favorieten"); setSelectedItem(null); }}>
          {t.tabFavorites} {favorites.length > 0 ? `(${favorites.length})` : ""}
        </button>
        <button className={`hf-tab ${activeTab === "interlock" ? "active" : ""}`} onClick={() => { setActiveTab("interlock"); setSelectedItem(null); }}>
          {t.tabInterlock}
        </button>
        <button className={`hf-tab ${activeTab === "testdoorn" ? "active" : ""}`} onClick={() => { setActiveTab("testdoorn"); setSelectedItem(null); }}>
          {t.tabMandrel}
        </button>
        <button className={`hf-tab ${activeTab === "onepiece" ? "active" : ""}`} onClick={() => { setActiveTab("onepiece"); setSelectedItem(null); }}>
          {t.tabOnePiece}
        </button>
        <button className={`hf-tab ${activeTab === "stagecrimp" ? "active" : ""}`} onClick={() => { setActiveTab("stagecrimp"); setSelectedItem(null); }}>
          {t.tabStageCrimp}
        </button>
        <button className={`hf-tab ${activeTab === "stainless" ? "active" : ""}`} onClick={() => { setActiveTab("stainless"); setSelectedItem(null); }}>
          {t.tabStainless}
        </button>
      </nav>

      <main className="screen-only px-4 py-4 max-w-6xl mx-auto">
        {activeTab === "zoeken" && (
          <>
            {/* Filterpaneel */}
            <section className="mb-4 p-3 sm:p-4 rounded flex flex-col gap-3" style={{ background: "var(--hf-bg-light)", border: "1px solid var(--hf-border)" }}>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div className="relative">
                  <label className="hf-subhead block mb-1">{t.filterSeriesLabel}</label>
                  <div className="flex items-center gap-2 bg-white rounded px-2" style={{ border: "1px solid var(--hf-border)" }}>
                    <Search size={16} style={{ color: "#999999" }} />
                    <input
                      value={seriesQuery}
                      onChange={(e) => { setSeriesQuery(e.target.value); setSelectedSeries(null); setShowSuggestions(true); }}
                      onFocus={() => setShowSuggestions(true)}
                      placeholder={t.filterSeriesPlaceholder}
                      className="w-full py-2.5 sm:py-2 text-sm outline-none"
                      style={{ background: "transparent" }}
                    />
                  </div>
                  {showSuggestions && seriesSuggestions.length > 0 && (
                    <div className="absolute z-10 left-0 right-0 mt-1 bg-white rounded shadow-lg" style={{ border: "1px solid var(--hf-border)" }}>
                      {seriesSuggestions.map((s) => (
                        <button
                          key={s}
                          onClick={() => selectSeries(s)}
                          className="w-full text-left px-3 py-2 text-sm hover:bg-gray-50"
                          style={{ color: "#000000" }}
                        >
                          {s}
                        </button>
                      ))}
                    </div>
                  )}
                </div>

                <div>
                  <label className="hf-subhead block mb-1">{t.filterMfrLabel}</label>
                  <div className="flex flex-wrap gap-1.5">
                    <button
                      onClick={() => setMfrFilter("")}
                      className={`hf-chip px-3 py-1.5 rounded ${mfrFilter === "" ? "active" : ""}`}
                    >
                      {t.filterAll}
                    </button>
                    {availableMfrCodes.map((c) => (
                      <button
                        key={c}
                        onClick={() => setMfrFilter(mfrFilter === c ? "" : c)}
                        className={`hf-chip px-3 py-1.5 rounded ${mfrFilter === c ? "active" : ""}`}
                      >
                        {c}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="hf-subhead block mb-1">{t.filterDnLabel}</label>
                  <div className="flex flex-wrap gap-1.5">
                    <button onClick={() => setDnFilter("")} className={`hf-chip px-3 py-1.5 rounded ${dnFilter === "" ? "active" : ""}`}>
                      {t.filterAll}
                    </button>
                    {availableDns.map((dn) => (
                      <button
                        key={dn}
                        onClick={() => setDnFilter(dnFilter === String(dn) ? "" : String(dn))}
                        className={`hf-chip px-3 py-1.5 rounded ${dnFilter === String(dn) ? "active" : ""}`}
                      >
                        DN {dn}
                      </button>
                    ))}
                  </div>
                </div>

                <div>
                  <label className="hf-subhead block mb-1">{t.filterVariantLabel}</label>
                  <div className="flex flex-wrap gap-1.5">
                    {[
                      { v: "alle", l: t.filterAll },
                      { v: "standaard", l: t.variantStandardChip },
                      { v: "geskived", l: t.variantSkivedChip },
                    ].map((o) => (
                      <button
                        key={o.v}
                        onClick={() => setVariantFilter(o.v)}
                        className={`hf-chip px-3 py-1.5 rounded ${variantFilter === o.v ? "active" : ""}`}
                      >
                        {o.l}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div className="flex items-center justify-between pt-1" style={{ borderTop: "1px solid var(--hf-border)" }}>
                <span className="hf-caption">{t.resultsCount(filteredResults.length)}</span>
                <div className="flex items-center gap-2">
                  <button onClick={openAddItem} className="hf-btn-outline rounded px-3 py-1.5 text-xs font-semibold flex items-center gap-1">
                    <Plus size={14} /> {t.btnAddCombination}
                  </button>
                  <button onClick={resetFilters} className="hf-btn-outline rounded px-3 py-1.5 text-xs font-semibold">
                    {t.btnClearFilters}
                  </button>
                </div>
              </div>
            </section>

            {filteredResults.length === 0 ? (
              <EmptyState title={t.emptyResultsTitle} message={t.emptyResultsMessage} />
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                {filteredResults.map((item) => (
                  <ResultCard
                    key={item.id}
                    item={item}
                    isFav={favorites.includes(item.id)}
                    onToggleFav={toggleFavorite}
                    onOpenDetail={setSelectedItem}
                    onEdit={openEditItem}
                    t={t}
                  />
                ))}
              </div>
            )}
          </>
        )}

        {activeTab === "favorieten" && (
          <>
            <div className="flex items-center justify-between mb-3">
              <span className="hf-subhead">{t.favoritesCount(favoriteItems.length)}</span>
            </div>
            {favoriteItems.length === 0 ? (
              <EmptyState title={t.emptyFavoritesTitle} message={t.emptyFavoritesMessage} />
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
                {favoriteItems.map((item) => (
                  <ResultCard
                    key={item.id}
                    item={item}
                    isFav={true}
                    onToggleFav={toggleFavorite}
                    onOpenDetail={setSelectedItem}
                    onEdit={openEditItem}
                    t={t}
                  />
                ))}
              </div>
            )}
          </>
        )}

        {activeTab === "testdoorn" && <TestMandrelPage t={t} />}
        {activeTab === "onepiece" && <OnePieceFittingPage t={t} />}
        {activeTab === "interlock" && <InterlockFittingPage t={t} />}
        {activeTab === "stagecrimp" && <StageCrimpPage t={t} exampleItem={activeData.find((d) => d.assemblyType === "stage")} />}
        {activeTab === "stainless" && <StainlessPage t={t} />}
      </main>

      {selectedItem && (
        <DetailModal
          item={selectedItem}
          isFav={favorites.includes(selectedItem.id)}
          onToggleFav={toggleFavorite}
          onClose={() => setSelectedItem(null)}
          onPrint={handlePrint}
          onEdit={openEditItem}
          onOpenOnePiece={() => { setSelectedItem(null); setActiveTab("onepiece"); }}
          onOpenStageCrimp={() => { setSelectedItem(null); setActiveTab("stagecrimp"); }}
          onOpenStainlessInfo={() => { setSelectedItem(null); setActiveTab("stainless"); }}
          t={t}
          lang={lang}
        />
      )}

      {editingItem && (
        <EditModal
          item={editingItem}
          isNew={isNewItem}
          onClose={closeEdit}
          onSave={handleSaveEdit}
          onDelete={handleDeleteItem}
          saving={saving}
          t={t}
        />
      )}

      {showUpdatePanel && (
        <UpdatePanel
          onClose={() => setShowUpdatePanel(false)}
          onSave={handleSaveUploadedData}
          onReset={handleResetToDefault}
          hasCustomData={!!customData}
          meta={dataMeta}
          saving={saving}
          storageOk={storageOk}
          t={t}
          locale={locale}
          data={activeData}
        />
      )}

      {/* Printweergave (alleen zichtbaar bij afdrukken) */}
      <div className="print-only" style={{ padding: "24px" }}>
        <div className="flex items-center gap-3" style={{ marginBottom: "16px" }}>
          <Logo />
          <div>
            <div className="hf-headline" style={{ fontSize: "20px" }}>{t.printTitle}</div>
            <div className="hf-caption">
              {t.printMetaLine(printLabel, new Date().toLocaleDateString(locale))}
            </div>
          </div>
        </div>

        <table className="hf-table" style={{ width: "100%", borderCollapse: "collapse", fontSize: "12px" }}>
          <thead>
            <tr style={{ background: "var(--hf-red)", color: "#ffffff" }}>
              <th style={{ textAlign: "left", padding: "6px" }}>{t.printThSeries}</th>
              <th style={{ textAlign: "left", padding: "6px" }}>DN</th>
              <th style={{ textAlign: "left", padding: "6px" }}>{t.printThMfrCode}</th>
              <th style={{ textAlign: "left", padding: "6px" }}>{t.printThHose}</th>
              <th style={{ textAlign: "left", padding: "6px" }}>{t.labelCrimpO}</th>
              <th style={{ textAlign: "left", padding: "6px" }}>{t.printThMandrel}</th>
              <th style={{ textAlign: "left", padding: "6px" }}>{t.printThD1LF}</th>
              <th style={{ textAlign: "left", padding: "6px" }}>{t.printThReference}</th>
            </tr>
          </thead>
          <tbody>
            {printItems.map((item, idx) => (
              <tr key={item.id} style={{ background: idx % 2 === 1 ? "var(--hf-bg-light)" : "#ffffff" }}>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>{item.series} ({variantShortLabel(item.uitvoering, t)})</td>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>{item.dn}</td>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>{item.mfrCodes.join(", ")}</td>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>{item.ferrule}</td>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>
                  {item.assemblyType === "stage" && item.crimp2 != null
                    ? `${stageCrimpDiameters(item).join(" → ")} mm`
                    : `${item.crimp.toFixed(1)} mm ${item.tol || ""}`}
                </td>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>{item.mandrel || "—"}</td>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>{item.d1} / {item.lf}</td>
                <td style={{ padding: "6px", borderBottom: "1px solid var(--hf-border)" }}>{item.standard}, {t.revShort(item.revDate, item.revSeq)}</td>
              </tr>
            ))}
          </tbody>
        </table>

        {selectedItem && (
          <div style={{ marginTop: "16px" }}>
            <div className="hf-subhead" style={{ marginBottom: "6px" }}>{t.printInstructionsHeading}</div>
            <ol style={{ paddingLeft: "18px", fontSize: "12px" }}>
              {getSteps(selectedItem, t, lang).map((s, i) => (
                <li key={i} style={{ marginBottom: "3px" }}>{s}</li>
              ))}
            </ol>
          </div>
        )}

        <p className="hf-caption" style={{ marginTop: "16px" }}>
          <strong>{t.disclaimerStrong}</strong> {t.disclaimerRest}
        </p>
      </div>
    </div>
  );
}
