import React, { useState } from 'react';
import {
  FileDown,
  FileSpreadsheet,
  FileText,
  FileType2,
  Printer,
  Trash2,
} from 'lucide-react';
import { PaperOrientation, PaperSizeId } from '../types';

export interface PrintPageConfig {
  orientation: PaperOrientation;
  paperSize: PaperSizeId;
}

interface TableExportToolbarProps {
  selectedCount?: number;
  defaultOrientation?: PaperOrientation;
  defaultPaperSize?: PaperSizeId;
  onDeleteSelected?: () => void;
  onDownloadTemplate?: () => void;
  onExportXlsx: () => void;
  onExportXls: () => void;
  onExportWord: (cfg: PrintPageConfig) => void;
  onExportPdf: (cfg: PrintPageConfig) => void;
  onPrint: (cfg: PrintPageConfig) => void;
}

const PAPER_SIZES: PaperSizeId[] = ['A4', 'F4', 'Letter', 'Legal', 'A5'];

export function TableExportToolbar({
  selectedCount = 0,
  defaultOrientation = 'landscape',
  defaultPaperSize = 'A4',
  onDeleteSelected,
  onDownloadTemplate,
  onExportXlsx,
  onExportXls,
  onExportWord,
  onExportPdf,
  onPrint,
}: TableExportToolbarProps) {
  const [orientation, setOrientation] = useState<PaperOrientation>(defaultOrientation);
  const [paperSize, setPaperSize] = useState<PaperSizeId>(defaultPaperSize);

  const cfg: PrintPageConfig = { orientation, paperSize };

  return (
    <div className="flex flex-wrap items-center gap-1.5">
      {selectedCount > 0 && onDeleteSelected && (
        <button
          type="button"
          onClick={onDeleteSelected}
          className="px-2.5 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap"
        >
          <Trash2 className="w-3.5 h-3.5" />
          Hapus ({selectedCount})
        </button>
      )}
      {onDownloadTemplate && (
        <button
          type="button"
          onClick={onDownloadTemplate}
          className="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs font-medium inline-flex items-center gap-1 whitespace-nowrap"
        >
          <FileSpreadsheet className="w-3.5 h-3.5 text-emerald-600" />
          Template
        </button>
      )}
      <button
        type="button"
        onClick={onExportXlsx}
        className="px-2.5 py-1.5 rounded-lg border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap"
      >
        <FileSpreadsheet className="w-3.5 h-3.5" />
        XLSX
      </button>
      <button
        type="button"
        onClick={onExportXls}
        className="px-2.5 py-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap"
      >
        <FileDown className="w-3.5 h-3.5" />
        XLS
      </button>

      {/* Orientasi Portrait / Landscape & Ukuran Kertas untuk Word, PDF, Cetak */}
      <div
        className="inline-flex items-center rounded-lg border border-slate-200 bg-slate-100 p-0.5 text-[11px]"
        role="group"
        aria-label="Orientasi halaman cetak"
      >
        <button
          type="button"
          onClick={() => setOrientation('portrait')}
          className={`px-2 py-1 rounded-md font-semibold transition-colors whitespace-nowrap ${
            orientation === 'portrait'
              ? 'bg-white text-indigo-700 shadow-xs'
              : 'text-slate-600 hover:text-slate-900'
          }`}
          title="Atur halaman cetak Portrait (Tegak)"
        >
          Portrait
        </button>
        <button
          type="button"
          onClick={() => setOrientation('landscape')}
          className={`px-2 py-1 rounded-md font-semibold transition-colors whitespace-nowrap ${
            orientation === 'landscape'
              ? 'bg-white text-indigo-700 shadow-xs'
              : 'text-slate-600 hover:text-slate-900'
          }`}
          title="Atur halaman cetak Landscape (Mendatar)"
        >
          Landscape
        </button>
      </div>

      <select
        value={paperSize}
        onChange={(e) => setPaperSize(e.target.value as PaperSizeId)}
        aria-label="Ukuran kertas cetak"
        className="px-2 py-1.5 text-[11px] font-semibold border border-slate-200 rounded-lg bg-white text-slate-700 focus:outline-none focus:border-indigo-600"
      >
        {PAPER_SIZES.map((sz) => (
          <option key={sz} value={sz}>
            {sz}
          </option>
        ))}
      </select>

      <button
        type="button"
        onClick={() => onExportWord(cfg)}
        className="px-2.5 py-1.5 rounded-lg border border-sky-200 bg-sky-50 hover:bg-sky-100 text-sky-800 text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap"
      >
        <FileType2 className="w-3.5 h-3.5" />
        Word
      </button>
      <button
        type="button"
        onClick={() => onExportPdf(cfg)}
        className="px-2.5 py-1.5 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap"
      >
        <FileText className="w-3.5 h-3.5" />
        PDF
      </button>
      <button
        type="button"
        onClick={() => onPrint(cfg)}
        className="px-2.5 py-1.5 rounded-lg bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold inline-flex items-center gap-1 whitespace-nowrap"
      >
        <Printer className="w-3.5 h-3.5" />
        Cetak
      </button>
    </div>
  );
}
